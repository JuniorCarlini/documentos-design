# documentos-design

Guia de estilo visual padrão para documentos da Saldus (Word e Google Docs): relatórios, propostas,
contratos, ofícios, atas, políticas e comunicados. **O texto segue a ABNT** (NBR 14724:2024,
NBR 6024, NBR 10520:2023 e NBR 6023): margens "Normal" do Word (2,5/3 cm, sem encadernação), fonte 12, entrelinha 1,5, recuo de 1,25 cm,
seções numeradas, legendas de tabelas e figuras, citações e referências. O visual usa as mesmas cores (`#12424D` petróleo,
`#FFB745` âmbar, branco) e fontes (**Sora** nos títulos, **Poppins** no texto) da skill
[planilhas-design](https://github.com/JuniorCarlini/planilhas-design). Planilhas e documentos saem da mesma família.

## Arquivos

| Arquivo | Para quê |
|---|---|
| `SKILL.md` | O guia completo. É o arquivo principal: funciona como "skill" e como instrução para qualquer IA |
| `tokens.json` | Os mesmos valores (cores, fontes, tamanhos, margens) em formato de dados |
| `scripts/estilo_word.py` | Gera documentos `.docx` no padrão (Python + python-docx) |
| `scripts/estilo_docs.gs` | Aplica o estilo no Google Docs (Apps Script) |
| `assets/modelos/` | Modelos prontos para o Word: documento formatado, documento simples, manual e proposta |
| `assets/logo-saldus.svg` | Logo Saldus em vetor (mestre), letras creme, para fundo escuro |
| `assets/logo-saldus-escura.svg` | Logo em vetor com letras petróleo, para fundo branco |
| `assets/logo.png` / `assets/logo-escura.png` | As mesmas logos em PNG de alta resolução, usadas pelo script do Word |
| `assets/simbolo-s-contorno.svg` | Contorno do "S" usado na marca d'água opcional |

## Como usar em cada plataforma

**Baixar:** `git clone https://github.com/JuniorCarlini/documentos-design.git` (ou *Code → Download ZIP*).

**Claude Code:**
```bash
git clone https://github.com/JuniorCarlini/documentos-design.git ~/.claude/skills/documentos-design
```

**Claude.ai (app/web):** compacte a pasta em `.zip` e envie em *Configurações → Capacidades → Skills*.

**ChatGPT:** crie um GPT (ou Projeto), cole o `SKILL.md` nas instruções e anexe `scripts/estilo_word.py`
e as duas logos como arquivos. O ChatGPT consegue rodar o script e devolver o `.docx`.

**Gemini:** crie um Gem e cole o `SKILL.md` nas instruções.

**Sem IA:** abra o modelo do tipo certo em `assets/modelos/` (documento, manual ou proposta), salve com outro nome e escreva usando os estilos
*Título 1*, *Título 2* e *Normal*. Use o checklist da seção 13 do `SKILL.md` para conferir.

## Word com Python
```bash
pip install python-docx matplotlib
python scripts/estilo_word.py exemplo.docx              # documento de exemplo
python scripts/estilo_word.py --modelo manual modelo.docx   # modelos: documento, documento-simples, manual, proposta
```
Capa: `capa={"estilo": "inteira"}` (folha toda petróleo) ou `capa={"imagem": "capa.png"}` (capa feita no Canva, A4 2480×3508 px).
Instituição que exige Arial/Times e texto preto: `novo_documento(..., fonte="Arial", marca=False)`.
Para o Word mostrar as fontes certas, instale Sora, Poppins e Poppins Light (grátis em fonts.google.com).

## Google Docs em 1 minuto
1. No documento, abra **Extensões → Apps Script**, cole `scripts/estilo_docs.gs` e salve.
2. Recarregue o documento. Vai aparecer o menu **Estilo**, com *Aplicar estilo padrão*, *Inserir faixa
   de título* e *Inserir caixa de destaque*.
3. A logo e o número de página entram pelo menu **Inserir**.
