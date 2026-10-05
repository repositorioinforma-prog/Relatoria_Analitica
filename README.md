# Relatória Analítica — Instituto Informa

Aplicativo desktop com Mapa de Calor, Matriz de Correlação, Cruzamentos, Análise Residual e Tutoriais.

## Organização

| Caminho | Conteúdo |
| --- | --- |
| `main.py` | Entrada do aplicativo; mantém o comando `python main.py` |
| `app/main.py` | Janela principal e navegação |
| `app/pages/` | Telas das ferramentas e dos tutoriais |
| `app/ui/` | Componentes de interface compartilhados |
| `app/core/resources.py` | Caminhos de recursos, incluindo execução com PyInstaller |
| `assets/` | Ícones, logo, guias HTML e imagens dos tutoriais |
| `packaging/` | Configuração do executável Windows |
| `scripts/` | Atalhos de instalação, execução e compilação no Windows |
| `requirements.txt` | Dependências para executar |
| `requirements-dev.txt` | Dependências para gerar o executável |

## Executar no Windows

Extraia o ZIP inteiro e mantenha as pastas juntas. Com Python instalado e disponível no terminal:

1. Execute `scripts/instalar.bat` para criar o ambiente e instalar as dependências.
2. Execute `scripts/iniciar.bat` para abrir o aplicativo.

Alternativa pelo terminal, na raiz do projeto:

```bat
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe main.py
```

## Gerar o executável Windows

Execute `scripts/compilar.bat`, ou use:

```bat
.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean --distpath dist --workpath build "packaging\Processamentos Relatoria.spec"
```

A saída fica em `dist/Processamentos Relatoria/`. Distribua a pasta completa, incluindo `_internal`; o `.exe` isolado não contém todas as dependências. Compile no Windows para gerar o aplicativo Windows. As pastas `build/` e `dist/` são geradas durante a compilação.

## Recursos e manutenção

Os caminhos dos recursos são calculados a partir do projeto ou do pacote PyInstaller; não dependem da pasta atual do terminal. Os guias HTML mantêm sua estrutura para preservar as referências às imagens.

Mapa de Calor e seus mapas de fundo carregam recursos da internet. A lista de dependências não fixa versões; mantenha seu ambiente já testado, se disponível, pois este ZIP não contém as versões instaladas no seu computador.

A reorganização ajusta a estrutura, importações e caminhos. Os cálculos das ferramentas permanecem iguais aos da versão limpa testada anteriormente. Esta versão reorganizada foi verificada por sintaxe, importações locais e resolução de recursos, mas a interface e o executável ainda precisam de teste no Windows.
