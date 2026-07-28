# Crítico Musical — Gerador de Relatórios de Podcasts

Ferramenta em Python para geração automática de relatórios detalhados em formato Word (`.docx`) a partir de episódios de podcasts musicais e culturais, com suporte a transcrição de áudio, análise fonográfica e citações.

## Funcionalidades

- **Análise de episódios** a partir de URL (Spotify, YouTube, RSS e outros)
- **Geração de relatórios** individuais ou consolidados em `.docx` com layout profissional
- **Transcrição de áudio** via Whisper (opcional) com fallback textual
- **Enriquecimento de dados** com metadados do episódio, músicas citadas e recomendações culturais
- **Base de episódios conhecidos** pré-configurada com descrições, setlist e dicas editoriais
- **Modo interativo** no terminal e suporte a linha de comando com argumentos
- **Barra de progresso animada** em 4 estágios: coleta → download → análise → compilação

## Requisitos

- Python 3.9 ou superior
- [`python-docx`](https://python-docx.readthedocs.io/)

## Instalação

```bash
# Clone o repositório
git clone https://github.com/saudebelem/critico-musical.git
cd critico-musical

# Instale a dependência principal
pip install python-docx
```

> **Opcional para transcrição via Whisper:**
> ```bash
> pip install openai-whisper
> ```

## Uso

### Modo interativo (recomendado)

Execute o script sem argumentos para o menu guiado:

```bash
python "Relatórios - críticas musicais.py"
```

Ou use o atalho Windows:

```bat
"Relatórios - críticas musicais.bat"
```

### Linha de comando

```bash
python "Relatórios - críticas musicais.py" \
  --urls "https://open.spotify.com/episode/..." \
  --mode 1 \
  --outdir "./saida"
```

**Parâmetros**

| Parâmetro | Descrição |
|-----------|-----------|
| `--urls` | Uma ou mais URLs de episódios separadas por espaço |
| `--mode` | `1` = relatórios separados · `2` = relatório consolidado · `3` = ambos |
| `--outdir` | Pasta de destino dos `.docx` gerados (padrão: diretório atual) |

## Estrutura do projeto

```
critico-musical/
├── Relatórios - críticas musicais.py   # Script principal
├── Relatórios - críticas musicais.bat  # Atalho Windows com auto-instalação de dependências
├── .gitignore
└── README.md
```

> As pastas `mídias baixadas/` e `relatórios aprovados/` são geradas automaticamente em execução e não são versionadas.

## Podcasts suportados (base pré-configurada)

O script inclui um dicionário interno (`KNOWN_EPISODES`) com metadados detalhados de episódios já catalogados, incluindo:

- **Podcast ABFP** — Tributo a José Mojica Marins, Art Rock com André Frateschi
- **Discoteca Básica** — *Are You Experienced* (The Jimi Hendrix Experience, 1967)
- **Discord and Rhyme** — Black Sabbath (1970)
- **O Som do Vinil** — Céu
- **Vamos Falar Sobre Música** — Gorillaz: Plastic Beach
- **Sala de Música CBN** — edições variadas
- **Podcast Cultural** — The Day I Tried to Live / Soundgarden

## Licença

Uso interno — Antigravity Projetos.
