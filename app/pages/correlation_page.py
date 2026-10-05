from app.core.resources import resource_base, resource_path
import os
import json
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

from PySide6.QtCore import QUrl, QObject, Slot
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QFileDialog, QMessageBox, QMainWindow
)
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWebEngineCore import QWebEngineProfile, QWebEnginePage
from PySide6.QtWebChannel import QWebChannel

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from app.ui.topbar import TopBar

try:
    from scipy import stats as scipy_stats
except Exception:
    scipy_stats = None




def _canon(s: str) -> str:
    s = (s or "").strip().lower()
    s = re.sub(r"\s+", " ", s)
    s = (
        s.replace("á", "a").replace("à", "a").replace("â", "a").replace("ã", "a")
         .replace("é", "e").replace("ê", "e")
         .replace("í", "i")
         .replace("ó", "o").replace("ô", "o").replace("õ", "o")
         .replace("ú", "u").replace("ç", "c")
    )
    s = re.sub(r"[^a-z0-9]+", " ", s).strip()
    return s


def _is_peso_name(name: str) -> bool:
    return "peso" in _canon(name)


def _numeric_candidates(df: pd.DataFrame) -> List[str]:
    cols: List[str] = []
    n = len(df)

    for c in df.columns:
        s = df[c]
        if pd.api.types.is_numeric_dtype(s):
            cols.append(str(c))
            continue

        if pd.api.types.is_object_dtype(s) or pd.api.types.is_string_dtype(s):
            x = pd.to_numeric(s.astype(str).str.replace(",", ".", regex=False), errors="coerce")
            if x.notna().sum() >= max(3, int(0.2 * n)):
                cols.append(str(c))

    return sorted(cols, key=lambda x: str(x).lower())


def _parse_missing_list(text: str) -> Tuple[List[str], List[float]]:
    tokens = [t.strip() for t in (text or "").split(",") if t.strip()]
    missing_str: List[str] = []
    missing_num: List[float] = []

    for t in tokens:
        tt = t.replace(".", "").replace(",", ".") if re.match(r"^[0-9\.\,]+$", t) else t
        try:
            v = float(tt)
            missing_num.append(v)
        except Exception:
            missing_str.append(t.strip().lower())

    return missing_str, missing_num


def _looks_numeric_token(t: str) -> bool:
    if t is None:
        return False
    s = str(t).strip()
    if not s:
        return False
    return bool(re.match(r"^[\+\-]?\d+([\,\.]\d+)?$", s))


def _token_to_float(t: str) -> Optional[float]:
    try:
        s = str(t).strip().replace(",", ".")
        return float(s)
    except Exception:
        return None


def _weighted_pearson(x: np.ndarray, y: np.ndarray, w: np.ndarray) -> float:
    sw = float(np.sum(w))
    if not np.isfinite(sw) or sw <= 0:
        return float("nan")

    mx = float(np.sum(w * x) / sw)
    my = float(np.sum(w * y) / sw)

    dx = x - mx
    dy = y - my

    vx = float(np.sum(w * dx * dx) / sw)
    vy = float(np.sum(w * dy * dy) / sw)
    if vx <= 0 or vy <= 0:
        return float("nan")

    cov = float(np.sum(w * dx * dy) / sw)
    return float(cov / np.sqrt(vx * vy))


SIG_P_MIN = 0.0
SIG_P_MAX = 0.005


def _sig_keep_value(p: float) -> bool:
    return p is not None and np.isfinite(p) and (SIG_P_MIN <= float(p) <= SIG_P_MAX)


def _sig_style_rgb(p: float) -> Optional[Tuple[str, str]]:
    """
    Estilo da folha Sig. (2-tailed).
    Destaca apenas os p-values que passam no critério 0,000 a 0,005.
    """
    if not _sig_keep_value(p):
        return None
    return ("D9EAD3", "274E13")  # verde claro / verde escuro


def _two_tailed_p_from_r(r: float, n: float) -> float:
    """
    p-valor bicaudal a partir de r e N.
    Aqui foi usado N ponderado como soma dos pesos válidos do par.
    """
    if scipy_stats is None:
        return float("nan")

    if r is None or not np.isfinite(r):
        return float("nan")
    if n is None or not np.isfinite(n) or float(n) <= 2.0:
        return float("nan")

    rr = float(r)
    rr = max(min(rr, 0.999999999999), -0.999999999999)

    den = 1.0 - (rr * rr)
    if den <= 0:
        return 0.0

    t = rr * np.sqrt((float(n) - 2.0) / den)
    return float(2.0 * scipy_stats.t.sf(abs(t), df=float(n) - 2.0))


def _weighted_corr_and_sig_matrices(
    df: pd.DataFrame,
    cols: List[str],
    w: pd.Series,
    method: str
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Retorna:
      corr_df      -> matriz de correlação "bruta"
      sig_df       -> matriz de Sig. (2-tailed)
      filtered_df  -> matriz da correlação já filtrada pelo p-value
    """
    method = (method or "pearson").lower().strip()
    n = len(cols)

    mat = np.full((n, n), np.nan, dtype=float)
    pmat = np.full((n, n), np.nan, dtype=float)
    filtered = np.full((n, n), np.nan, dtype=float)

    wv = pd.to_numeric(w, errors="coerce").fillna(0.0).astype(float)

    for i in range(n):
        mat[i, i] = 1.0
        filtered[i, i] = 1.0
        pmat[i, i] = np.nan

    for i in range(n):
        xi = pd.to_numeric(df[cols[i]], errors="coerce")

        for j in range(i + 1, n):
            yj = pd.to_numeric(df[cols[j]], errors="coerce")

            mask = xi.notna() & yj.notna() & (wv > 0)
            if int(mask.sum()) < 2:
                continue

            ww = wv[mask].to_numpy(dtype=float)
            xx = xi[mask]
            yy = yj[mask]

            if method == "spearman":
                xx = xx.rank(method="average")
                yy = yy.rank(method="average")

            r = _weighted_pearson(xx.to_numpy(dtype=float), yy.to_numpy(dtype=float), ww)

            # N ponderado do par
            n_pair = float(np.sum(ww))
            p = _two_tailed_p_from_r(r, n_pair)

            mat[i, j] = r
            mat[j, i] = r

            pmat[i, j] = p
            pmat[j, i] = p

            if _sig_keep_value(p):
                filtered[i, j] = r
                filtered[j, i] = r

    return (
        pd.DataFrame(mat, index=cols, columns=cols),
        pd.DataFrame(pmat, index=cols, columns=cols),
        pd.DataFrame(filtered, index=cols, columns=cols),
    )


def _corr_style_rgb(v: float) -> Optional[Tuple[str, str]]:
    if v is None or not np.isfinite(v):
        return None
    if abs(float(v) - 1.0) < 1e-12:
        return None

    if v <= -0.5:
        return ("BE5014", "FFFFFF")
    if v <= -0.4:
        return ("F1A983", "BE5014")
    if v <= -0.3:
        return ("F7C7AC", "BE5014")
    if v <= -0.2:
        return ("FBE2D5", "BE5014")

    if v >= 0.5:
        return ("104861", "FFFFFF")
    if v >= 0.4:
        return ("44B3E1", "104861")
    if v >= 0.3:
        return ("83CCEB", "104861")
    if v >= 0.2:
        return ("C0E6F5", "104861")

    return None


def _format_token_for_ui(v: Any) -> Tuple[str, str]:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ("__NA__", "⟂ (NA)")
    if isinstance(v, str):
        s = v.strip()
        if s == "":
            return ("__EMPTY__", "<vazio>")
        return (s, s)

    try:
        if isinstance(v, (np.integer, int)):
            s = str(int(v))
            return (s, s)
        if isinstance(v, (np.floating, float)):
            if float(v).is_integer():
                s = str(int(v))
            else:
                s = f"{float(v):g}"
            return (s, s)
    except Exception:
        pass

    s = str(v).strip()
    if s == "":
        return ("__EMPTY__", "<vazio>")
    return (s, s)

def _jsonable_matrix(mat: np.ndarray) -> List[List[Optional[float]]]:
    out: List[List[Optional[float]]] = []
    arr = np.asarray(mat, dtype=float)

    for row in arr:
        new_row: List[Optional[float]] = []
        for v in row:
            if np.isfinite(v):
                new_row.append(float(v))
            else:
                new_row.append(None)
        out.append(new_row)

    return out

HTML = r"""<!doctype html>
<html lang="pt-br">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Matriz de Correlação</title>
  <script src="qrc:///qtwebchannel/qwebchannel.js"></script>

  <style>
    html,body{height:100%;margin:0;font-family:system-ui,Arial,sans-serif;}
    .app{display:grid;grid-template-columns:520px 1fr;height:100%}
    .panel{border-right:1px solid #eee;padding:14px 16px 16px;overflow:auto;background:#fff}
    h1{font-size:18px;margin:0 0 12px}
    h2{font-size:14px;margin:16px 0 8px}
    label{font-size:13px;display:block;margin:8px 0 6px}
    small{color:#666;font-size:12px}

    button,input,select,textarea{
      font:inherit;padding:8px 10px;border-radius:10px;border:1px solid #ddd;background:#fff
    }
    textarea{width:100%;box-sizing:border-box}
    .btn-primary{background:#1E90FF;border-color:#1E90FF;color:#fff}
    .btn-outline{background:#fff;color:#111;border-color:#ccc}
    .btn-sm{padding:6px 8px;border-radius:10px;font-size:12px}

    .inline{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
    .muted{color:#666;font-size:12px}
    .box{border:1px solid #eee;border-radius:14px;padding:10px 10px 8px;background:#fafafa;margin-bottom:10px;}

    .listbox{
      border:1px solid #e6e6e6;border-radius:12px;background:#fff;
      padding:6px; max-height:220px; overflow:auto;
    }
    .row{display:flex;align-items:center;gap:8px;padding:4px 4px}
    .row input[type="checkbox"]{margin:0}
    .row .lab{font-size:12px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; flex:1;}
    .row .cnt{font-size:12px;color:#666;min-width:70px;text-align:right;font-variant-numeric: tabular-nums;}

    .search{width:100%; box-sizing:border-box}

    /* Chips */
    .chips{display:flex;gap:6px;flex-wrap:wrap;margin-top:8px}
    .chip{
      display:inline-flex;align-items:center;gap:6px;
      border:1px solid #e6e6e6;background:#fff;border-radius:999px;
      padding:4px 8px;font-size:12px;max-width:100%;
    }
    .chip .x{cursor:pointer;font-weight:800;color:#444}

    /* Renomear list */
    .renameList{border:1px solid #e6e6e6;border-radius:12px;background:#fff;padding:8px;max-height:220px;overflow:auto;}
    .renameItem{display:grid;grid-template-columns:90px 1fr;gap:8px;align-items:center;padding:6px 2px;border-bottom:1px solid #f2f2f2}
    .renameItem:last-child{border-bottom:none}
    .renameKey{font-size:12px;color:#333;font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
    .renameInp{width:100%;box-sizing:border-box;font-size:12px}

    /* Área da tabela */
    #main{height:100%;background:#fff;overflow:hidden;display:flex;flex-direction:column;}
    #headerbar{padding:10px 12px;border-bottom:1px solid #eee;display:flex;align-items:center;justify-content:space-between;gap:10px;}
    #title{font-weight:800;font-size:14px;color:#111;margin:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
    #subtitle{font-size:12px;color:#666;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
    #tableWrap{flex:1;overflow:auto;padding:12px;}

    table{border-collapse:separate;border-spacing:0;font-size:12px;min-width:720px;}
    th, td{border:1px solid #e6e6e6;padding:6px 8px;background:#fff;text-align:center;font-variant-numeric: tabular-nums;}
    th{font-weight:800;position:sticky;top:0;z-index:5;background:#f7f7f7;}
    td.firstcol, th.firstcol{position:sticky;left:0;z-index:6;background:#f7f7f7;font-weight:800;text-align:left;}
    td.val{min-width:78px;}
    td.blank{background:#fff;color:#fff;}

    .badge{font-size:12px;padding:4px 8px;border-radius:999px;border:1px solid #e6e6e6;background:#fff;white-space:nowrap;}
  </style>
</head>

<body>
<div class="app">
  <aside class="panel">
    <h1>Matriz de Correlação</h1>

    <div class="box">
      <h2 style="margin:0 0 8px">0) Dados</h2>
      <label>Importar Banco</label>
      <div class="inline">
        <button id="btnPickData" class="btn-outline" type="button">Importar CSV/XLSX/SAV</button>
        <small id="dataStatus" class="muted"></small>
      </div>
    </div>

    <div class="box">
      <h2 style="margin:0 0 8px">1) Peso (obrigatório)</h2>
      <label>Variável de peso</label>
      <select id="weightVar"></select>
      <small class="muted" id="weightHint">Mostra primeiro colunas que contêm “peso”.</small>
    </div>

    <div class="box">
      <h2 style="margin:0 0 8px">2) Missing (por variável selecionada)</h2>
      <small class="muted">Aqui aparecem SOMENTE as variáveis que você marcou em “Selecionar colunas”.</small>

      <label style="margin-top:10px">Variável (selecionada)</label>
      <div class="inline">
        <select id="missingVar" style="flex:1;min-width:220px"></select>
        <button id="btnLoadVals" class="btn-outline btn-sm" type="button">Carregar valores</button>
      </div>

      <div class="inline" style="margin-top:8px;justify-content:space-between">
        <div class="inline">
          <button id="btnMissAll" class="btn-outline btn-sm" type="button">Todos</button>
          <button id="btnMissNone" class="btn-outline btn-sm" type="button">Nenhum</button>
          <button id="btnAddMissing" class="btn-primary btn-sm" type="button">Adicionar missing</button>
          <button id="btnClearMissingVar" class="btn-outline btn-sm" type="button">Limpar desta variável</button>
        </div>
        <span class="badge" id="missMeta">—</span>
      </div>

      <input id="missingSearch" class="search" type="text" placeholder="Buscar valor..." style="margin-top:10px"/>
      <div id="missingValuesList" class="listbox" style="margin-top:8px"></div>

      <label style="margin-top:10px">Missing aplicado nesta variável</label>
      <div id="missingChips" class="chips"></div>

      <label style="margin-top:10px">Missing manual (opcional, global)</label>
      <input id="missingVals" type="text" placeholder="Ex.: 98, 99, Não sei, Branco/Nulo" />
      <small class="muted">Esse campo é um “extra” (aplica em todas as variáveis).</small>
    </div>

    <div class="box">
      <h2 style="margin:0 0 8px">3) Nomes (edição dinâmica)</h2>
      <small class="muted">Você pode renomear aqui OU dar 2 cliques no cabeçalho da tabela (estilo Excel).</small>

      <div class="inline" style="margin-top:8px">
        <label style="display:flex;gap:8px;align-items:center;margin:0;font-size:12px;color:#333">
          <input id="chkUseSpssLabels" type="checkbox" checked />
          Usar rótulos do SPSS (inicial)
        </label>
      </div>

      <label style="margin-top:10px">Variáveis selecionadas</label>
      <div id="renameList" class="renameList"></div>
    </div>

    <div class="box">
      <h2 style="margin:0 0 8px">4) Configurações</h2>

      <label>Método</label>
      <select id="method">
        <option value="pearson">Pearson</option>
        <option value="spearman">Spearman</option>
        <option value="kendall">Kendall</option>
      </select>
      <small class="muted">Obs.: Kendall com peso vira “sem ponderação”.</small>

      <label style="margin-top:10px">Casas decimais</label>
      <input id="decimals" type="number" min="0" max="10" step="1" value="3" style="width:120px"/>

      <div class="inline" style="margin-top:8px">
        <label style="display:flex;gap:8px;align-items:center;margin:0;font-size:12px;color:#333">
          <input id="chkConvert" type="checkbox" checked />
          Tentar converter texto em número
        </label>
      </div>

      <div class="inline" style="margin-top:10px">
        <button id="btnCalc" class="btn-primary" type="button">Calcular</button>
        <small id="calcStatus" class="muted"></small>
      </div>
    </div>

    <div class="box">
      <h2 style="margin:0 0 8px">5) Selecionar colunas</h2>
      <small class="muted">Selecione as variáveis numéricas para compor a matriz.</small>

      <div class="inline" style="margin-top:10px; justify-content:space-between">
        <div class="inline">
          <button id="btnAll" class="btn-outline btn-sm" type="button">Todos</button>
          <button id="btnNone" class="btn-outline btn-sm" type="button">Nenhum</button>
        </div>
        <span class="badge" id="badgeCount">0 selecionadas</span>
      </div>

      <input id="searchCols" class="search" type="text" placeholder="Buscar..." style="margin-top:10px"/>
      <div id="colsList" class="listbox" style="margin-top:8px"></div>
    </div>

    <div class="box">
      <h2 style="margin:0 0 8px">6) Exportar</h2>
      <div class="inline" style="margin-top:8px">
        <button id="btnExportCsv" class="btn-outline" type="button">Exportar CSV</button>
        <button id="btnExportXlsx" class="btn-primary" type="button">Exportar XLSX</button>
        <small id="expStatus" class="muted"></small>
      </div>
      <small class="muted">XLSX sai com col=10, linha=40, fonte DIN 10 e cores do VBA.</small>
    </div>
  </aside>

  <main id="main">
    <div id="headerbar">
      <div style="min-width:0">
        <div id="title">Matriz de correlação</div>
        <div id="subtitle">Carregue uma base e clique em Calcular.</div>
      </div>
      <span class="badge" id="badgeMeta">—</span>
    </div>
    <div id="tableWrap">
      <div class="muted">Nada para mostrar ainda.</div>
    </div>
  </main>
</div>

<script>
  window.QtBridge = null;
  new QWebChannel(qt.webChannelTransport, function(channel){
    window.QtBridge = channel.objects.QtBridge;
  });

  function el(id){ return document.getElementById(id); }

  // colunas numéricas (candidatas)
  let __allCols = [];
  let __selected = new Set();

  // labels base (SPSS / inicial) por variável original
  let __baseLabel = {};   // var -> label inicial (ou var)
  // label atual editável
  let __dispMap = {};     // var -> label atual

  let __weights = [];

  // correlação em memória
  let __corr = null;      // payload do python
  let __vars = [];        // vars na ordem do cálculo (originais)
  let __matrix = null;

  // Missing guiado
  let __missingRules = {};      // var -> Set(tokens)
  let __uniqueCache = {};       // var -> payload de únicos
  let __uniqueSelected = new Set();

  function dispVar(varName){
    return (__dispMap && __dispMap[varName]) ? __dispMap[varName] : ( (__baseLabel[varName]) ? __baseLabel[varName] : varName );
  }

  function updateBadge(){
    el('badgeCount').textContent = `${__selected.size} selecionadas`;
  }

  function renderCols(){
    const box = el('colsList');
    box.innerHTML = '';
    const q = (el('searchCols').value || '').toLowerCase().trim();

    const show = __allCols.filter(c => {
      const dn = String(dispVar(c)).toLowerCase();
      const vn = String(c).toLowerCase();
      return !q || dn.includes(q) || vn.includes(q);
    });

    for(const c of show){
      const row = document.createElement('div');
      row.className = 'row';

      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.checked = __selected.has(c);
      cb.addEventListener('change', ()=>{
        if(cb.checked) __selected.add(c);
        else __selected.delete(c);
        updateBadge();
        refreshSelectedDependentUI();
      });

      const lab = document.createElement('div');
      lab.className = 'lab';
      lab.textContent = dispVar(c);
      lab.title = c;

      row.appendChild(cb);
      row.appendChild(lab);
      box.appendChild(row);
    }
  }

  function renderWeights(){
    const sel = el('weightVar');
    sel.innerHTML = '';

    const opt0 = document.createElement('option');
    opt0.value = '';
    opt0.textContent = '— selecione o peso —';
    sel.appendChild(opt0);

    for(const w of (__weights || [])){
      const op = document.createElement('option');
      op.value = w;
      op.textContent = dispVar(w);
      sel.appendChild(op);
    }

    if(!(__weights || []).length){
      for(const c of (__allCols || [])){
        const op = document.createElement('option');
        op.value = c;
        op.textContent = dispVar(c);
        sel.appendChild(op);
      }
      el('weightHint').textContent = 'Não achei colunas com “peso”. Mostrando todas para você escolher.';
    } else {
      el('weightHint').textContent = 'Mostra primeiro colunas que contêm “peso”.';
    }
  }

  function selectedArray(){
    return Array.from(__selected.values());
  }

  // ---------------- Missing UI (somente selecionadas)
  function renderMissingVarOptions(){
    const sel = el('missingVar');
    const prev = sel.value || '';

    const sels = selectedArray();
    sel.innerHTML = '';

    if(!sels.length){
      const op = document.createElement('option');
      op.value = '';
      op.textContent = '— selecione colunas abaixo —';
      sel.appendChild(op);

      el('missingValuesList').innerHTML = '<small class="muted">Selecione colunas em “Selecionar colunas”.</small>';
      el('missMeta').textContent = '—';
      el('missingChips').innerHTML = '<small class="muted">—</small>';
      return;
    }

    for(const v of sels){
      const op = document.createElement('option');
      op.value = v;
      op.textContent = dispVar(v);
      sel.appendChild(op);
    }

    if(prev && sels.includes(prev)) sel.value = prev;
    else sel.value = sels[0];

    renderMissingChips();
    renderUniqueValues(sel.value || '');
  }

  function renderMissingChips(){
    const varName = el('missingVar').value || '';
    const box = el('missingChips');
    box.innerHTML = '';
    if(!varName){
      box.innerHTML = '<small class="muted">—</small>';
      return;
    }

    const set = (__missingRules[varName] || new Set());
    if(!set.size){
      box.innerHTML = '<small class="muted">Nenhum missing configurado.</small>';
      return;
    }

    for(const token of Array.from(set.values())){
      const chip = document.createElement('div');
      chip.className = 'chip';
      chip.title = token;

      const label = document.createElement('span');
      label.textContent = token === '__EMPTY__' ? '<vazio>' : token;

      const x = document.createElement('span');
      x.className = 'x';
      x.textContent = '×';
      x.addEventListener('click', ()=>{
        set.delete(token);
        __missingRules[varName] = set;
        renderMissingChips();
      });

      chip.appendChild(label);
      chip.appendChild(x);
      box.appendChild(chip);
    }
  }

  function renderUniqueValues(varName){
    const list = el('missingValuesList');
    list.innerHTML = '';
    __uniqueSelected = new Set();

    const payload = __uniqueCache[varName];
    if(!varName){
      list.innerHTML = '<small class="muted">—</small>';
      el('missMeta').textContent = '—';
      return;
    }

    if(!payload || !payload.ok){
      list.innerHTML = '<small class="muted">Clique em “Carregar valores”.</small>';
      el('missMeta').textContent = '—';
      return;
    }

    const q = (el('missingSearch').value || '').toLowerCase().trim();
    const values = (payload.values || []).filter(v=>{
      const d = String(v.display || '').toLowerCase();
      const t = String(v.token || '').toLowerCase();
      return !q || d.includes(q) || t.includes(q);
    });

    el('missMeta').textContent = `${payload.n_unique} únicos · mostrando ${values.length}`;

    if(!values.length){
      list.innerHTML = '<small class="muted">Nada encontrado.</small>';
      return;
    }

    for(const v of values){
      const row = document.createElement('div');
      row.className = 'row';

      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.addEventListener('change', ()=>{
        if(cb.checked) __uniqueSelected.add(v.token);
        else __uniqueSelected.delete(v.token);
      });

      const lab = document.createElement('div');
      lab.className = 'lab';
      lab.textContent = v.display;
      lab.title = v.token;

      const cnt = document.createElement('div');
      cnt.className = 'cnt';
      cnt.textContent = String(v.count);

      row.appendChild(cb);
      row.appendChild(lab);
      row.appendChild(cnt);
      list.appendChild(row);
    }
  }

  function loadUniqueValues(force=false){
    const varName = el('missingVar').value || '';
    if(!varName) return;

    if(!force && __uniqueCache[varName] && __uniqueCache[varName].ok){
      renderUniqueValues(varName);
      return;
    }

    el('missingValuesList').innerHTML = '<small class="muted">Carregando valores...</small>';
    el('missMeta').textContent = 'carregando...';

    const req = { col: varName, max: 200 };
    window.QtBridge?.getUniqueValues?.(JSON.stringify(req));
  }

  el('missingSearch').addEventListener('input', ()=>{
    const varName = el('missingVar').value || '';
    renderUniqueValues(varName);
  });

  el('missingVar').addEventListener('change', ()=>{
    renderMissingChips();
    renderUniqueValues(el('missingVar').value || '');
  });

  el('btnLoadVals').addEventListener('click', ()=> loadUniqueValues(true));

  el('btnMissAll').addEventListener('click', ()=>{
    const varName = el('missingVar').value || '';
    const payload = __uniqueCache[varName];
    if(!payload || !payload.ok) return;

    __uniqueSelected = new Set((payload.values || []).map(v=>v.token));
    renderUniqueValues(varName);
    document.querySelectorAll('#missingValuesList input[type="checkbox"]').forEach(cb=> cb.checked = true);
  });

  el('btnMissNone').addEventListener('click', ()=>{
    __uniqueSelected = new Set();
    document.querySelectorAll('#missingValuesList input[type="checkbox"]').forEach(cb=> cb.checked = false);
  });

  el('btnAddMissing').addEventListener('click', ()=>{
    const varName = el('missingVar').value || '';
    if(!varName) return;
    if(!__uniqueSelected.size) return;

    const set = (__missingRules[varName] || new Set());
    for(const t of __uniqueSelected.values()){
      if(t === '__NA__') continue;
      set.add(t);
    }
    __missingRules[varName] = set;
    renderMissingChips();
  });

  el('btnClearMissingVar').addEventListener('click', ()=>{
    const varName = el('missingVar').value || '';
    if(!varName) return;
    __missingRules[varName] = new Set();
    renderMissingChips();
  });

  // ---------------- Renomear (somente selecionadas)
  function updateLabelInDom(varName){
    const label = dispVar(varName);
    document.querySelectorAll(`[data-var="${CSS.escape(varName)}"]`).forEach(node=>{
      // se estiver editando (tem input), não sobrescrever
      if(node.querySelector && node.querySelector('input')) return;
      node.textContent = label;
    });

    // também atualiza <option> dos selects (missingVar / weightVar) sem re-render geral
    // (prático e rápido)
    Array.from(el('missingVar').options).forEach(op=>{
      if(op.value === varName) op.textContent = label;
    });
  }

  function renderRenameList(){
    const box = el('renameList');
    box.innerHTML = '';

    const sels = selectedArray();
    if(!sels.length){
      box.innerHTML = '<small class="muted">Selecione colunas primeiro.</small>';
      return;
    }

    for(const v of sels){
      const row = document.createElement('div');
      row.className = 'renameItem';

      const key = document.createElement('div');
      key.className = 'renameKey';
      key.textContent = v;
      key.title = v;

      const inp = document.createElement('input');
      inp.className = 'renameInp';
      inp.type = 'text';
      inp.value = dispVar(v);
      inp.placeholder = v;

      inp.addEventListener('input', ()=>{
        const nv = (inp.value || '').trim();
        __dispMap[v] = nv ? nv : (__baseLabel[v] || v);
        updateLabelInDom(v);
      });

      row.appendChild(key);
      row.appendChild(inp);
      box.appendChild(row);
    }
  }

  // chamado quando muda seleção (colunas) -> atualiza MissingVar e Renomear list
  function refreshSelectedDependentUI(){
    renderMissingVarOptions();
    renderRenameList();
  }

  // ---------------- Tabela
  function fmt(v, decimals){
    if(v === null || v === undefined || !isFinite(v)) return '—';
    const d = Math.max(0, Math.min(10, Number(decimals)||3));
    return Number(v).toFixed(d);
  }

  function rgb(r,g,b){ return `rgb(${r},${g},${b})`; }

  function corrStyle(v){
    if(v === null || v === undefined || !isFinite(v)) return null;
    if(Math.abs(v - 1) < 1e-12) return null;

    if(v <= -0.5) return { bg: rgb(190, 80, 20),  fg: rgb(255,255,255) };
    if(v <= -0.4) return { bg: rgb(241,169,131), fg: rgb(190, 80, 20) };
    if(v <= -0.3) return { bg: rgb(247,199,172), fg: rgb(190, 80, 20) };
    if(v <= -0.2) return { bg: rgb(251,226,213), fg: rgb(190, 80, 20) };

    if(v >=  0.5) return { bg: rgb(16, 72, 97),  fg: rgb(255,255,255) };
    if(v >=  0.4) return { bg: rgb(68,179,225), fg: rgb(16, 72, 97) };
    if(v >=  0.3) return { bg: rgb(131,204,235), fg: rgb(16, 72, 97) };
    if(v >=  0.2) return { bg: rgb(192,230,245), fg: rgb(16, 72, 97) };

    return null;
  }

  function startInlineEdit(node, varName){
    if(!node) return;
    if(node.querySelector && node.querySelector('input')) return;

    const old = dispVar(varName);
    node.textContent = '';

    const inp = document.createElement('input');
    inp.type = 'text';
    inp.value = old;
    inp.style.width = '100%';
    inp.style.boxSizing = 'border-box';
    inp.style.fontSize = '12px';
    inp.style.padding = '4px 6px';
    inp.style.borderRadius = '8px';
    inp.style.border = '1px solid #ccc';

    const commit = (apply=true)=>{
      const nv = (inp.value || '').trim();
      if(apply){
        __dispMap[varName] = nv ? nv : (__baseLabel[varName] || varName);
      }
      node.textContent = dispVar(varName);
      updateLabelInDom(varName);
      // re-render renameList sem perder tudo? (não!)
      // só atualiza o input correspondente se existir
      document.querySelectorAll('.renameItem').forEach(item=>{
        const key = item.querySelector('.renameKey');
        const input = item.querySelector('input');
        if(key && input && key.textContent === varName){
          input.value = dispVar(varName);
        }
      });
    };

    inp.addEventListener('keydown', (e)=>{
      if(e.key === 'Enter'){ e.preventDefault(); inp.blur(); }
      if(e.key === 'Escape'){ e.preventDefault(); commit(false); }
    });
    inp.addEventListener('blur', ()=> commit(true));

    node.appendChild(inp);
    inp.focus();
    inp.select();
  }

  function renderTable(payload){
    const wrap = el('tableWrap');
    wrap.innerHTML = '';

    if(!payload || !payload.vars || !payload.matrix){
      wrap.innerHTML = '<div class="muted">Nada para mostrar.</div>';
      return;
    }

    const vars = payload.vars;
    const mat = payload.matrix;
    const decimals = payload.decimals;

    __vars = vars.slice();
    __matrix = mat;

    const table = document.createElement('table');

    const thead = document.createElement('thead');
    const hr = document.createElement('tr');

    const th0 = document.createElement('th');
    th0.className = 'firstcol';
    th0.textContent = '';
    hr.appendChild(th0);

    vars.forEach(v=>{
      const th = document.createElement('th');
      th.textContent = dispVar(v);
      th.setAttribute('data-var', v);
      th.title = 'Duplo clique para renomear';
      th.addEventListener('dblclick', ()=> startInlineEdit(th, v));
      hr.appendChild(th);
    });

    thead.appendChild(hr);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');

    for(let i=0;i<vars.length;i++){
      const tr = document.createElement('tr');

      const td0 = document.createElement('td');
      td0.className = 'firstcol';
      td0.textContent = dispVar(vars[i]);
      td0.setAttribute('data-var', vars[i]);
      td0.title = 'Duplo clique para renomear';
      td0.addEventListener('dblclick', ()=> startInlineEdit(td0, vars[i]));
      tr.appendChild(td0);

      for(let j=0;j<vars.length;j++){
        const td = document.createElement('td');
        td.className = 'val';

        if(i > j){
          td.textContent = '';
          td.classList.add('blank');
          td.style.background = '#fff';
        } else {
          const v = mat[i][j];
          td.textContent = fmt(v, decimals);

          const st = corrStyle(v);
          if(st){
            td.style.background = st.bg;
            td.style.color = st.fg;
            td.style.fontWeight = '700';
          } else {
            td.style.background = '#fff';
            td.style.color = '#111';
            td.style.fontWeight = '400';
          }
        }
        tr.appendChild(td);
      }

      tbody.appendChild(tr);
    }

    table.appendChild(tbody);
    wrap.appendChild(table);
  }

  // ---------------- Eventos gerais
  el('searchCols').addEventListener('input', renderCols);

  el('btnAll').addEventListener('click', ()=>{
    __allCols.forEach(c=> __selected.add(c));
    updateBadge();
    renderCols();
    refreshSelectedDependentUI();
  });

  el('btnNone').addEventListener('click', ()=>{
    __selected = new Set();
    updateBadge();
    renderCols();
    refreshSelectedDependentUI();
  });

  // Buttons
  el('btnPickData').addEventListener('click', ()=> window.QtBridge?.pickData?.());

  el('btnCalc').addEventListener('click', ()=>{
    const method = el('method').value || 'pearson';
    const decimals = Math.max(0, Math.min(10, Number(el('decimals').value)||3));
    const convert = !!el('chkConvert').checked;

    const weightVar = el('weightVar').value || '';
    if(!weightVar){
      el('calcStatus').textContent = 'Selecione o peso (obrigatório).';
      return;
    }

    const cols = selectedArray();
    if(cols.length < 2){
      el('calcStatus').textContent = 'Selecione pelo menos 2 colunas.';
      return;
    }

    // missing rules: var -> array tokens
    const missingRulesObj = {};
    for(const k of Object.keys(__missingRules)){
      missingRulesObj[k] = Array.from(__missingRules[k].values());
    }

    const payload = {
      cols,
      weightVar,
      missingVals: el('missingVals').value || '',
      missingRules: missingRulesObj,
      useSpssLabels: !!el('chkUseSpssLabels').checked,
      method,
      decimals,
      convertText: convert,
    };

    el('calcStatus').textContent = 'Calculando...';
    window.QtBridge?.computeCorr?.(JSON.stringify(payload));
  });

  el('btnExportCsv').addEventListener('click', ()=>{
    if(!__vars.length || !__matrix){
      el('expStatus').textContent = 'Calcule antes de exportar.';
      return;
    }
    const names = __vars.map(v => dispVar(v));
    window.QtBridge?.exportCorrCsv?.(JSON.stringify({ display: names }));
  });

  el('btnExportXlsx').addEventListener('click', ()=>{
    if(!__vars.length || !__matrix){
      el('expStatus').textContent = 'Calcule antes de exportar.';
      return;
    }
    const names = __vars.map(v => dispVar(v));
    window.QtBridge?.exportCorrXlsx?.(JSON.stringify({ display: names }));
  });

  // ---------------- Python -> JS
  window.__setColumns = function(payloadJson){
    try{
      const data = JSON.parse(payloadJson);
      el('dataStatus').textContent = data.label || '';

      __allCols = data.columns || [];
      __weights = data.weight_candidates || [];

      // base labels
      __baseLabel = {};
      __dispMap = {};
      const labels = data.labels || {};
      for(const c of __allCols){
        const lab = (labels[c] || '').trim();
        __baseLabel[c] = lab ? lab : c;
        __dispMap[c] = __baseLabel[c];
      }

      // pré-seleciona até 12
      __selected = new Set(__allCols.slice(0, Math.min(__allCols.length, 12)));
      updateBadge();
      renderCols();
      renderWeights();

      // reset missing
      __missingRules = {};
      __uniqueCache = {};
      __uniqueSelected = new Set();
      el('missingSearch').value = '';
      el('missingVals').value = '';
      el('missingValuesList').innerHTML = '<small class="muted">Selecione uma variável (selecionada) e clique em “Carregar valores”.</small>';
      el('missMeta').textContent = '—';

      refreshSelectedDependentUI();

      // reset tabela
      __corr = null;
      __vars = [];
      __matrix = null;

      el('calcStatus').textContent = '';
      el('title').textContent = 'Matriz de correlação';
      el('subtitle').textContent = 'Selecione colunas, missing e clique em Calcular.';
      el('badgeMeta').textContent = '—';
      el('tableWrap').innerHTML = '<div class="muted">Pronto para calcular.</div>';

    }catch(e){ console.error(e); }
  };

  window.__setUniqueValues = function(payloadJson){
    try{
      const p = JSON.parse(payloadJson);
      __uniqueCache[p.col] = p;
      renderUniqueValues(p.col);
    }catch(e){
      console.error(e);
      el('missingValuesList').innerHTML = '<small class="muted">Erro ao carregar valores.</small>';
      el('missMeta').textContent = 'erro';
    }
  };

  window.__setCorrResult = function(payloadJson){
    try{
      const r = JSON.parse(payloadJson);
      if(!r.ok){
        el('calcStatus').textContent = r.error || 'Erro ao calcular.';
        return;
      }
      __corr = r;

      // vars originais
      const vars = r.vars || [];
      const initialDisplay = r.cols_display || [];

      // aplica labels iniciais retornados no cálculo (ex.: SPSS labels)
      for(let i=0;i<vars.length;i++){
        const v = vars[i];
        const lab = (initialDisplay[i] || '').trim();
        if(lab){
          __baseLabel[v] = lab;
          // só sobrescreve o dispMap se ainda estava “igual base antiga”
          __dispMap[v] = __dispMap[v] ? __dispMap[v] : lab;
          // se usuário não editou nada ainda, fica igual base
          __dispMap[v] = (__dispMap[v] === v || __dispMap[v] === (__baseLabel[v] || v)) ? lab : __dispMap[v];
        } else {
          __baseLabel[v] = __baseLabel[v] || v;
          __dispMap[v] = __dispMap[v] || __baseLabel[v];
        }
      }

      el('calcStatus').textContent = 'Pronto.';
      el('title').textContent = r.title || 'Matriz de correlação';
      el('subtitle').textContent = r.subtitle || '';
      el('badgeMeta').textContent = `${(r.method||'').toUpperCase()} · ${vars.length} vars`;

      renderRenameList();
      renderTable(r);

    }catch(e){
      console.error(e);
      el('calcStatus').textContent = 'Erro ao receber resultado.';
    }
  };
</script>
</body>
</html>
"""

class CorrelationBridge(QObject):
    def __init__(self, main_window: QMainWindow, web_view: QWebEngineView):
        super().__init__()
        self.main_window = main_window
        self.web_view = web_view

        self.df: Optional[pd.DataFrame] = None
        self.data_label: str = ""
        self.var_labels: Dict[str, str] = {}

        # último resultado
        self.last_cols_vars: List[str] = []
        self.last_matrix: Optional[np.ndarray] = None          # matriz filtrada (exibida/exportada)
        self.last_raw_matrix: Optional[np.ndarray] = None      # matriz bruta
        self.last_p_matrix: Optional[np.ndarray] = None        # Sig. (2-tailed)
        self.last_method: str = "pearson"
        self.last_decimals: int = 3
        self.last_weight_var: str = ""

    def _push_columns(self):
        if self.df is None:
            payload = {"columns": [], "label": "", "labels": {}, "weight_candidates": []}
        else:
            cols = _numeric_candidates(self.df)
            labels = {c: (self.var_labels.get(c) or "") for c in cols}
            weights = [str(c) for c in self.df.columns if _is_peso_name(str(c))]
            payload = {
                "columns": cols,
                "label": self.data_label,
                "labels": labels,
                "weight_candidates": weights,
            }

        self.web_view.page().runJavaScript(
            f"window.__setColumns({json.dumps(json.dumps(payload, ensure_ascii=False))});"
        )

    @Slot()
    def pickData(self):
        path, _ = QFileDialog.getOpenFileName(
            self.main_window,
            "Selecionar base de dados",
            "",
            "Dados (*.csv *.xlsx *.sav);;Todos (*.*)"
        )
        if not path:
            return

        try:
            ext = os.path.splitext(path)[1].lower()
            self.var_labels = {}

            if ext == ".csv":
                try:
                    df = pd.read_csv(path, encoding="utf-8")
                except UnicodeDecodeError:
                    df = pd.read_csv(path, encoding="latin1")

            elif ext == ".xlsx":
                df = pd.read_excel(path)

            elif ext == ".sav":
                try:
                    import pyreadstat  # type: ignore
                except Exception:
                    raise RuntimeError("Para ler .sav, instale: pip install pyreadstat")

                df = None
                meta = None
                last_err = None

                file_enc = None
                try:
                    _, m0 = pyreadstat.read_sav(path, metadataonly=True)  # type: ignore
                    file_enc = (getattr(m0, "file_encoding", None) or "").strip() or None
                except Exception:
                    file_enc = None

                candidates = []
                for enc in [None, file_enc, "utf-8", "cp1252", "latin1", "iso-8859-1", "iso-8859-15", "utf-16le", "ucs-2le"]:
                    if enc and enc not in candidates:
                        candidates.append(enc)

                for enc in candidates:
                    try:
                        if enc is None:
                            df, meta = pyreadstat.read_sav(path)  # type: ignore
                        else:
                            df, meta = pyreadstat.read_sav(path, encoding=enc)  # type: ignore
                        break
                    except Exception as e:
                        last_err = e
                        msg = str(e)
                        if (
                            "Unable to convert string" in msg
                            or "invalid byte sequence" in msg
                            or "codec can't decode" in msg
                            or "invalid continuation byte" in msg
                        ):
                            continue
                        raise

                if df is None or meta is None:
                    raise last_err or Exception("Falha ao ler o arquivo .sav.")

                labels_map: Dict[str, str] = {}
                try:
                    m = getattr(meta, "column_names_to_labels", None)
                    if isinstance(m, dict):
                        labels_map = {str(k): str(v) for k, v in m.items() if v}
                    else:
                        cn = getattr(meta, "column_names", None)
                        cl = getattr(meta, "column_labels", None)
                        if cn and cl and len(cn) == len(cl):
                            labels_map = {str(k): str(v) for k, v in zip(cn, cl) if v}
                except Exception:
                    labels_map = {}

                self.var_labels = labels_map

            else:
                raise RuntimeError("Formato não suportado. Use CSV, XLSX ou SAV.")

            df = df.dropna(axis=1, how="all")
            self.df = df
            self.data_label = os.path.basename(path)

            # reset último resultado
            self.last_cols_vars = []
            self.last_matrix = None
            self.last_raw_matrix = None
            self.last_p_matrix = None
            self.last_method = "pearson"
            self.last_decimals = 3
            self.last_weight_var = ""

            self._push_columns()

        except Exception as e:
            QMessageBox.critical(self.main_window, "Erro", f"Falha ao carregar dados.\n\n{e}")

    @Slot(str)
    def getUniqueValues(self, req_json: str):
        try:
            if self.df is None:
                raise RuntimeError("Nenhum arquivo carregado.")

            req = json.loads(req_json) if req_json else {}
            col = str(req.get("col") or "").strip()
            max_items = int(req.get("max") or 200)
            max_items = max(20, min(500, max_items))

            if not col or col not in self.df.columns:
                raise RuntimeError("Coluna inválida.")

            s = self.df[col]

            if pd.api.types.is_numeric_dtype(s):
                vc = s.dropna().value_counts().head(max_items)
                items = []
                for val, cnt in vc.items():
                    token, disp = _format_token_for_ui(val)
                    if token == "__NA__":
                        continue
                    items.append({"token": token, "display": disp, "count": int(cnt)})
            else:
                ss = s.copy()
                mask = ss.notna()
                ss2 = ss[mask].astype(str).map(lambda x: x.strip())
                ss2 = ss2.replace("", "__EMPTY__")
                vc = ss2.value_counts().head(max_items)
                items = []
                for val, cnt in vc.items():
                    if val == "__EMPTY__":
                        items.append({"token": "__EMPTY__", "display": "<vazio>", "count": int(cnt)})
                    else:
                        items.append({"token": str(val), "display": str(val), "count": int(cnt)})

            payload = {
                "ok": True,
                "col": col,
                "n_unique": int(s.dropna().nunique()) if int(s.notna().sum()) else 0,
                "values": items,
            }
        except Exception as e:
            col = ""
            try:
                col = str(json.loads(req_json).get("col", "")) if req_json else ""
            except Exception:
                col = ""
            payload = {"ok": False, "col": col, "error": str(e), "values": [], "n_unique": 0}

        self.web_view.page().runJavaScript(
            f"window.__setUniqueValues({json.dumps(json.dumps(payload, ensure_ascii=False))});"
        )

    @Slot(str)
    def computeCorr(self, payload_json: str):
        try:
            if self.df is None:
                raise RuntimeError("Nenhum arquivo carregado.")

            payload = json.loads(payload_json) if payload_json else {}
            if not isinstance(payload, dict):
                raise RuntimeError("Payload inválido.")

            cols = payload.get("cols") or []
            weight_var = (payload.get("weightVar") or "").strip()

            missing_vals_global = payload.get("missingVals") or ""
            missing_rules = payload.get("missingRules") or {}

            use_spss_labels = bool(payload.get("useSpssLabels", True))
            method = (payload.get("method") or "pearson").lower().strip()
            decimals = int(payload.get("decimals") or 3)
            convert_text = bool(payload.get("convertText", True))

            if not isinstance(cols, list) or len(cols) < 2:
                raise RuntimeError("Selecione pelo menos 2 colunas.")
            cols = [str(c) for c in cols]

            for c in cols:
                if c not in self.df.columns:
                    raise RuntimeError(f"Coluna inválida: {c}")

            if not weight_var:
                raise RuntimeError("Selecione o peso (obrigatório).")
            if weight_var not in self.df.columns:
                raise RuntimeError("Variável de peso inválida.")

            if method not in ("pearson", "spearman", "kendall"):
                method = "pearson"

            if method in ("pearson", "spearman") and scipy_stats is None:
                raise RuntimeError("Para calcular Sig. (2-tailed), instale: pip install scipy")

            df2 = self.df[cols + [weight_var]].copy()

            # 1) Missing guiado por variável
            if isinstance(missing_rules, dict):
                for c, tokens in missing_rules.items():
                    c = str(c)
                    if c not in df2.columns:
                        continue
                    if not isinstance(tokens, list):
                        continue

                    ser = df2[c]

                    if "__EMPTY__" in tokens:
                        try:
                            if pd.api.types.is_string_dtype(ser) or pd.api.types.is_object_dtype(ser):
                                empty_mask = ser.notna() & ser.astype(str).map(lambda x: x.strip() == "")
                                df2.loc[empty_mask, c] = np.nan
                        except Exception:
                            pass

                    num_tokens = [t for t in tokens if _looks_numeric_token(t)]
                    str_tokens = [str(t).strip().lower() for t in tokens if not _looks_numeric_token(t) and t not in ("__EMPTY__", "__NA__")]

                    if num_tokens:
                        tmp_num = pd.to_numeric(
                            ser.astype(str).str.replace(",", ".", regex=False),
                            errors="coerce"
                        )
                        for t in num_tokens:
                            fv = _token_to_float(t)
                            if fv is None:
                                continue
                            df2.loc[tmp_num == fv, c] = np.nan

                    if str_tokens:
                        tmp_txt = ser.astype(str).str.strip().str.lower()
                        df2.loc[tmp_txt.isin(str_tokens), c] = np.nan

            # 2) Missing global manual
            missing_str, missing_num = _parse_missing_list(str(missing_vals_global))
            if missing_str:
                for c in cols:
                    tmp_txt = df2[c].astype(str).str.strip().str.lower()
                    df2.loc[tmp_txt.isin(missing_str), c] = np.nan

            # 3) conversão + missing numérico global
            for c in cols:
                if convert_text:
                    df2[c] = pd.to_numeric(
                        df2[c].astype(str).str.replace(",", ".", regex=False),
                        errors="coerce"
                    )
                else:
                    df2[c] = pd.to_numeric(df2[c], errors="coerce")

                if missing_num:
                    for mv in missing_num:
                        df2.loc[df2[c] == mv, c] = np.nan

            # peso
            w = pd.to_numeric(df2[weight_var], errors="coerce").fillna(0.0).astype(float)

            valid_w = w > 0
            dfv = df2.loc[valid_w, cols].copy()
            wv = w.loc[valid_w].copy()

            # labels iniciais
            cols_display: List[str] = []
            for c in cols:
                if use_spss_labels:
                    lab = (self.var_labels.get(c) or "").strip()
                    cols_display.append(lab if lab else c)
                else:
                    cols_display.append(c)

            # cálculo
            if method == "kendall":
                corr_df = dfv.corr(method="kendall")
                raw_mat = corr_df.to_numpy(dtype=float)
                p_mat = np.full_like(raw_mat, np.nan, dtype=float)
                filtered_mat = raw_mat.copy()
                method_used = "kendall"
                subtitle_extra = f"· peso: {weight_var} (Kendall sem ponderação)"
            else:
                corr_df, sig_df, filtered_df = _weighted_corr_and_sig_matrices(dfv, cols, wv, method=method)
                raw_mat = corr_df.to_numpy(dtype=float)
                p_mat = sig_df.to_numpy(dtype=float)
                filtered_mat = filtered_df.to_numpy(dtype=float)
                method_used = method
                subtitle_extra = f"· peso: {weight_var} · filtro Sig. 0,000 a 0,005"

            # guarda para export
            self.last_cols_vars = cols
            self.last_raw_matrix = raw_mat.copy()
            self.last_p_matrix = p_mat.copy()
            self.last_matrix = filtered_mat.copy()   # matriz exibida/exportada
            self.last_method = method_used
            self.last_decimals = int(decimals)
            self.last_weight_var = weight_var

            payload_out = {
              "ok": True,
              "title": "Matriz de correlação",
              "subtitle": f"{self.data_label} · {len(cols)} variáveis {subtitle_extra}",
              "method": method_used,
              "decimals": int(decimals),
              "vars": cols,
              "cols_display": cols_display,
              "matrix": _jsonable_matrix(filtered_mat),
            }

            self.web_view.page().runJavaScript(
                f"window.__setCorrResult({json.dumps(json.dumps(payload_out, ensure_ascii=False))});"
            )

        except Exception as e:
            payload_out = {"ok": False, "error": str(e)}
            self.web_view.page().runJavaScript(
                f"window.__setCorrResult({json.dumps(json.dumps(payload_out, ensure_ascii=False))});"
            )

    @Slot(str)
    def exportCorrCsv(self, payload_json: str):
        try:
            if self.last_matrix is None or not self.last_cols_vars:
                QMessageBox.information(self.main_window, "Exportar", "Calcule a correlação antes de exportar.")
                return

            payload = json.loads(payload_json) if payload_json else {}
            display = payload.get("display") or []
            if not isinstance(display, list) or len(display) != len(self.last_cols_vars):
                display = self.last_cols_vars[:]  # fallback

            out_path, _ = QFileDialog.getSaveFileName(
                self.main_window,
                "Salvar CSV",
                "correlacao.csv",
                "CSV (*.csv)"
            )
            if not out_path:
                return
            if not out_path.lower().endswith(".csv"):
                out_path += ".csv"

            df = pd.DataFrame(self.last_matrix, index=display, columns=display)

            for i in range(df.shape[0]):
                for j in range(df.shape[1]):
                    if i > j:
                        df.iat[i, j] = np.nan

            df = df.round(int(self.last_decimals))
            df.to_csv(out_path, index=True, sep=";", encoding="utf-8-sig")
            QMessageBox.information(self.main_window, "Exportar", "CSV salvo com sucesso.")

        except Exception as e:
            QMessageBox.critical(self.main_window, "Erro", f"Falha ao exportar CSV.\n\n{e}")

    @Slot(str)
    def exportCorrXlsx(self, payload_json: str):
        try:
            if self.last_matrix is None or not self.last_cols_vars:
                QMessageBox.information(self.main_window, "Exportar", "Calcule a correlação antes de exportar.")
                return

            payload = json.loads(payload_json) if payload_json else {}
            display = payload.get("display") or []
            if not isinstance(display, list) or len(display) != len(self.last_cols_vars):
                display = self.last_cols_vars[:]  # fallback

            out_path, _ = QFileDialog.getSaveFileName(
                self.main_window,
                "Salvar XLSX",
                "correlacao.xlsx",
                "Excel (*.xlsx)"
            )
            if not out_path:
                return
            if not out_path.lower().endswith(".xlsx"):
                out_path += ".xlsx"

            cols = display
            mat = self.last_matrix
            pmat = self.last_p_matrix
            dec = int(self.last_decimals)

            wb = Workbook()
            ws = wb.active
            ws.title = "Correlações"

            font_base = Font(name="DIN", size=10, bold=False, color="000000")
            font_head = Font(name="DIN", size=10, bold=True, color="000000")
            align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
            align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

            thin = Side(style="thin", color="D9D9D9")
            border = Border(left=thin, right=thin, top=thin, bottom=thin)

            def write_matrix_sheet(
                ws_obj,
                data_matrix: np.ndarray,
                *,
                blank_diag: bool,
                number_format: str,
                style_func
            ):
                n = len(cols)

                for c in range(1, n + 2):
                    ws_obj.column_dimensions[get_column_letter(c)].width = 10
                for r in range(1, n + 2):
                    ws_obj.row_dimensions[r].height = 40

                ws_obj.cell(1, 1, "").font = font_head
                ws_obj.cell(1, 1).alignment = align_left
                ws_obj.cell(1, 1).border = border

                for j, name in enumerate(cols, start=2):
                    cell = ws_obj.cell(1, j, name)
                    cell.font = font_head
                    cell.alignment = align_center
                    cell.border = border

                for i, row_name in enumerate(cols, start=2):
                    c0 = ws_obj.cell(i, 1, row_name)
                    c0.font = font_head
                    c0.alignment = align_left
                    c0.border = border

                    for j in range(2, n + 2):
                        ii = i - 2
                        jj = j - 2

                        cell = ws_obj.cell(i, j)
                        cell.border = border
                        cell.alignment = align_center
                        cell.font = font_base

                        if ii > jj:
                            cell.value = None
                            continue

                        if blank_diag and ii == jj:
                            cell.value = None
                            continue

                        v = float(data_matrix[ii, jj]) if np.isfinite(data_matrix[ii, jj]) else np.nan
                        if not np.isfinite(v):
                            cell.value = None
                            continue

                        cell.value = v
                        cell.number_format = number_format

                        st = style_func(v) if style_func is not None else None
                        if st is not None:
                            fill_hex, font_hex = st
                            cell.fill = PatternFill("solid", fgColor=fill_hex)
                            cell.font = Font(name="DIN", size=10, bold=True, color=font_hex)

                ws_obj.freeze_panes = "B2"

            corr_num_fmt = "0" if dec <= 0 else ("0." + ("0" * dec))
            write_matrix_sheet(
                ws,
                mat,
                blank_diag=False,
                number_format=corr_num_fmt,
                style_func=_corr_style_rgb
            )

            ws_sig = wb.create_sheet("Sig. (2-tailed)")
            if pmat is None:
                pmat = np.full((len(cols), len(cols)), np.nan, dtype=float)

            write_matrix_sheet(
                ws_sig,
                pmat,
                blank_diag=True,
                number_format="0.000",
                style_func=_sig_style_rgb
            )

            wb.save(out_path)
            QMessageBox.information(self.main_window, "Exportar", "XLSX salvo com sucesso.")

        except Exception as e:
            QMessageBox.critical(self.main_window, "Erro", f"Falha ao exportar XLSX.\n\n{e}")


class CorrelationPage(QWidget):
    def __init__(self, main_window: QMainWindow, profile: Optional[QWebEngineProfile] = None):
        super().__init__()
        self.main_window = main_window

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.topbar = TopBar(
            title="Funcionalidade: Matriz de Correlação",
            icon_path=resource_path("assets/icons/correlation.ico"),
            back_icon_path=resource_path("assets/icons/arrow-left.ico"),
            show_help=True
        )
        
        self.btn_back = self.topbar.btn_back
        layout.addWidget(self.topbar)

        self.view = QWebEngineView(self)
        layout.addWidget(self.view, 1)

        if profile is None:
            profile = QWebEngineProfile.defaultProfile()
        profile.downloadRequested.connect(lambda d: d.cancel())
        page = QWebEnginePage(profile, self.view)
        self.view.setPage(page)

        self.view.setHtml(HTML, QUrl("https://app.local/"))

        self.bridge = CorrelationBridge(main_window, self.view)
        channel = QWebChannel(self.view.page())
        channel.registerObject("QtBridge", self.bridge)
        self.view.page().setWebChannel(channel)

    def cleanup(self):
        pass