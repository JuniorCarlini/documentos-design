---
name: documentos-design
description: Guia de estilo padrão (identidade Saldus + normas ABNT) para documentos de texto — Word (.docx) e Google Docs — inclusive manuais com capturas de tela marcadas (números e setas), capas e página de aviso — com texto no padrão ABNT (NBR 14724:2024, NBR 6024, NBR 10520:2023, NBR 6023) e as cores e a logo da Saldus na faixa de título, tabelas, cartões, caixas de destaque, gráficos e assinaturas. Use SEMPRE que for criar, formatar ou "deixar bonito" qualquer documento da empresa — relatório, proposta comercial, contrato, ofício, carta, ata de reunião, política, procedimento, manual, passo a passo, tutorial, memorando, comunicado, trabalho ou texto técnico — mesmo que o usuário não peça um estilo específico, para que todos os documentos saiam com a mesma identidade visual e no padrão ABNT.
---

# Design de Documentos — Padrão Saldus + ABNT

Todos os documentos de texto da Saldus seguem **duas camadas**:

1. **O texto segue a ABNT**, na versão mais recente (sem a margem de encadernação): fonte 12, entrelinha 1,5, recuo de
   parágrafo, seções numeradas, citações, legendas de tabelas e figuras, e referências.
2. **O visual segue a marca Saldus:** faixa petróleo com a logo, títulos em petróleo, tabelas iguais
   às das planilhas, cartões e caixas de destaque. As cores e fontes são as mesmas da skill
   `planilhas-design`.

Siga os valores exatos abaixo em vez de improvisar. Se o usuário pedir algo diferente, atenda o
pedido dele; fora isso, use este padrão.

**Atalho:** em Python, use `scripts/estilo_word.py`, que já aplica tudo. Ele está descrito na seção 10. Antes de entregar, siga a seção 11.

**Normas usadas:** ABNT NBR 14724:2024 (apresentação, versão corrigida em 2025), NBR 6024
(numeração progressiva das seções), NBR 10520:2023 (citações) e NBR 6023 (referências). Se o
documento for para uma instituição com manual próprio (faculdade, órgão público), o manual dela vale
mais que este guia.

---

## 0. Antes de começar: pergunte

Antes de montar o documento, faça **de uma vez só** as perguntas abaixo que o pedido ainda não
responde. Use opções de múltipla escolha, com a recomendada primeiro, para o usuário só clicar. Na
mesma mensagem, diga em uma ou duas frases **o que você vai fazer** e o que fica de fora. Assim a
pessoa confirma antes de você gastar tempo. Não pergunte o que já foi dito nem o que tem resposta
óbvia, e faça no máximo 4 perguntas por vez. Se faltarem mais, junte as de menor peso (aviso e marca
d'água) numa só.

1. **Tipo de documento:** relatório, proposta comercial, manual / passo a passo, contrato, ata,
   política / procedimento, comunicado, ofício / carta.
2. **Primeira página** (para documento formal, não pergunte: é sempre só o título sóbrio, veja a seção 4):
   - faixa de título (recomendado para até 4 páginas e para todo documento formal)
   - capa em bloco
   - capa inteira em petróleo (com ou sem **foto** no alto: peça a imagem)
   - capa com imagem feita fora (peça o arquivo A4, 2480 × 3508 px)
3. **Página de aviso:**
   - nenhuma (recomendado)
   - confidencial
   - uso interno
   - minuta para revisão
   - texto próprio
4. **Cabeçalho, rodapé e marca d'água** (são sempre opcionais: pergunte toda vez, a não ser que o
   usuário já tenha dito; se ele não responder, use o padrão recomendado):
   - **sem cabeçalho + rodapé em faixa petróleo** com a logo à esquerda e o número da página à direita
     (recomendado e padrão da Saldus)
   - rodapé com linha fina e "Saldus" no lugar da faixa
   - sem número de página
   - cabeçalho no alto (logo, número e filete), **só se a pessoa pedir**
   - e, à parte: sem marca d'água (recomendado para relatórios e propostas)
   - com o **"S" da Saldus** em contorno bege bem claro no fundo de todas as páginas, menos na capa
     (fica bem em manuais e documentos confidenciais)
5. **Conteúdo extra** (pode marcar várias):
   - cartões com os números principais
   - tabela com total
   - gráfico
   - citações e referências ABNT
   - assinaturas
   - **capturas de tela marcadas** (manual: peça os prints)

Para manual, peça também **os prints de cada passo** e o caminho que o leitor faz, por exemplo
"Rural > Orientações-Fomento". Para proposta e contrato, peça os dados da **ficha**: cliente, data,
validade e as partes.

Exemplo de resumo antes das perguntas: *"Vou montar um manual de 3 passos com capa inteira, prints
marcados com números e uma caixa de dicas. Não vou incluir referências nem assinaturas. Confirma as
opções abaixo?"*

## 1. Cores

**Marca:** `#12424D` (petróleo) · `#FFB745` (âmbar) · `#FFFFFF` (branco).

| Token | Hex | Uso no documento |
|---|---|---|
| `primaria` | `#12424D` | Faixa do título, títulos de seção, cabeçalho de tabela, palavras em **negrito**, "Tabela 1 –" |
| `destaque` | `#FFB745` | Filetes (sob a faixa, sob o cabeçalho da página, sob o cabeçalho da tabela), barra lateral das caixas, o cartão do número principal e o **destaque do título da capa** (sobre fundo petróleo). **Nunca como cor de texto sobre fundo branco** |
| `primaria-clara` | `#E7EFF0` | Fundo das caixas de destaque, dos cartões e da linha de total |
| `zebra` | `#F5F8F8` | Linhas alternadas das tabelas |
| `borda` | `#D5E0E2` | Linhas finas: entre linhas de tabela e acima do rodapé (nunca sob os títulos) |
| `texto` | `#1E2B2E` | Todo o texto corrido (quase preto, como pede a ABNT) |
| `texto-suave` | `#5F7378` | Rótulos, número da página, rodapé |
| `texto-faixa` | `#B7CDD2` | Subtítulo dentro da faixa petróleo |
| `positivo` / `negativo` / `alerta` | `#2E7D32` / `#C62828` / `#B26A00` | Status e valores bons, ruins e pendentes nas tabelas |
| `entrada` | `#FFF1D6` | Fundo da caixa de "Atenção" |

**Sobre cor e ABNT:** a norma recomenda texto em preto e reserva cores para as ilustrações. Por isso
o texto corrido fica em `#1E2B2E` (quase preto) e a cor da marca aparece só nos títulos, tabelas,
caixas e faixa. Se o documento for para uma instituição que exige tudo em preto, use `marca=False`
(veja a seção 10).

**Regra do âmbar:** é tempero. Use em filetes, numa barra lateral, em **um** cartão, na faixa de aviso e
no destaque do título da capa (texto âmbar só sobre o petróleo).

## 2. Fontes e tamanhos (ABNT: fonte 12 no texto, menor em citações, notas, legendas e tabelas)

| Elemento | Fonte | Tamanho | Estilo |
|---|---|---|---|
| Texto corrido | Poppins | **12** | `texto`, **justificado**, entrelinha **1,5**, recuo de **1,25 cm** na 1ª linha, sem espaço entre parágrafos |
| Seção primária: `1 INTRODUÇÃO` | **Sora** | 12 | **MAIÚSCULAS, negrito**, petróleo, sem linha embaixo |
| Seção secundária: `1.1 Objetivo` | **Sora** | 12 | **negrito**, petróleo |
| Seção terciária: `1.1.1 Escopo` | Poppins | 12 | **negrito**, `texto` |
| Seção quaternária: `1.1.1.1 Item` | Poppins | 12 | *itálico*, `texto` |
| Citação longa (mais de 3 linhas) | Poppins | **10** | recuo de **4 cm**, entrelinha simples, justificada, sem aspas |
| Tabelas, legendas e fontes | Poppins (cabeçalho em Sora) | **10** | entrelinha simples |
| Referências | Poppins | 12 | entrelinha simples, alinhadas à esquerda, uma linha em branco entre elas |
| Faixa de título | Sora | 18 | negrito, branco, com subtítulo Poppins 10 itálico `texto-faixa` |
| Título da capa inteira | Poppins **Light** + **Sora** negrito | 34 | uma frase só: parte fina branca + destaque grosso âmbar |
| Faixa de aviso ("ATENÇÃO") | Sora | 18 / 8,5 | negrito, petróleo, sobre âmbar |

A ABNT pede **tamanho 12**, mas não fixa a família da fonte. Por isso mantemos Sora e Poppins, da
marca. Instale **Sora, Poppins (Regular, Bold, SemiBold, Italic) e Poppins Light** em quem for abrir no
Word. Se a instituição exigir Arial ou Times, use `fonte="Arial"` (seção 10). Sem Sora/Poppins
instaladas (gratuitas no Google Fonts), o Word troca as fontes. O idioma do documento é
**português (Brasil)**.

## 3. Página (ABNT)

- **A4**, margens **"Normal" do Word: superior e inferior 2,5 cm, esquerda e direita 3 cm**. Largura
  útil: **15 cm**. Não use a margem de encadernação da ABNT (3/3/2/2), porque os documentos da Saldus são
  digitais ou impressos soltos. Use 3/3/2/2 só se pedirem um trabalho impresso e encadernado.
- **Padrão Saldus: SEM cabeçalho.** A logo e o número da página ficam no **rodapé em faixa** petróleo
  (veja abaixo). Não coloque cabeçalho a menos que o usuário peça. A ABNT sugere a paginação no alto,
  mas o padrão da empresa é no rodapé.
- A **1ª página conta, mas não mostra o número**, porque a faixa já leva a logo. A capa também não mostra.
- **Cabeçalho (só se pedirem):** `cabecalho=True` mostra logo, número e filete âmbar no alto, e
  `cabecalho="numero"` mostra só o número. Aí o número sai do rodapé. A marca d'água funciona com ou sem
  cabeçalho.
- **Rodapé em faixa (padrão)**: faixa petróleo de 1,5 cm, de ponta a ponta no pé da
  folha, com a logo clara à esquerda (alinhada à margem do texto) e o número da página em branco, Sora
  negrito, à direita. O número só aparece se não estiver no cabeçalho; `paginacao=False` tira o número.
- **Rodapé em linha (alternativa):** `estilo_rodape="linha"` mostra "Saldus" (ou "Saldus · Confidencial")
  em 8 pt `texto-suave`, com linha fina em cima e o número da página à direita.
- Títulos ficam presos ao parágrafo seguinte, para nunca sobrar título sozinho no pé da página.

## 4. Abertura: faixa de título ou capa

**Documento formal NUNCA tem capa.** Isso vale para política, contrato, termo, procedimento, norma,
regulamento, ata, ofício, comunicado oficial e consentimento. É assim que grandes empresas fazem: a 1ª
página tem só o **título sóbrio** e, na linha cinza logo abaixo dele, assunto · Versão · Vigência ·
Classificação. **Sem quadro/ficha de controle.** O texto começa na mesma página. Cláusulas numeradas
(**1.1**, **1.2**) com o número em negrito. Capa só em material de apresentação: relatório longo,
proposta comercial, manual, apresentação institucional.

**Faixa de título (padrão):** um bloco petróleo no topo da 1ª página, com a largura exata do texto.
Tem o título em branco, o subtítulo embaixo (assunto, data, área) e a **logo clara**
(`assets/logo.png`, 3,2 cm) à direita. A logo-mestre está em vetor em `assets/logo-saldus.svg` (e
`logo-saldus-escura.svg` para fundo branco); os PNG são gerados dela em alta resolução. Um filete âmbar de 3 pt fecha a faixa.

**Capa (documentos longos, com 5 páginas ou mais, ou propostas formais):** a 1ª página vira só a
capa. Ela nunca tem número de página nem cabeçalho, e o conteúdo começa na página 2. Há três estilos:

| Estilo | Como fica | Quando usar |
|---|---|---|
| `bloco` (padrão) | Folha branca com um bloco petróleo alto no topo (logo em cima, título Sora 28 embaixo, filete âmbar) e a **ficha** (Preparado para, Data, Versão) abaixo | Relatórios e políticas |
| `inteira` | **Folha toda em petróleo**, de ponta a ponta, com faixa âmbar no pé e logo clara no alto. O **título, numa frase só com parte fina branca e destaque grosso em âmbar** (34 pt), o subtítulo curto e a ficha formam um bloco centralizado na área petróleo | Propostas, manuais e apresentações para clientes |
| `inteira` + **foto** | A mesma capa, mas com uma **foto na parte de cima**, de ponta a ponta e cortada sem distorcer. A altura se ajusta: a foto ocupa tudo o que o texto não usa (de 11 a 20 cm) e um filete âmbar embaixo dela. Logo, título e ficha ficam na área petróleo | Quando há uma imagem principal: produto, equipe, evento |
| `imagem` | **Capa pronta feita fora** (Canva, Figma, Photoshop) com a imagem principal escolhida. A imagem ocupa a folha inteira | Materiais especiais, quando a capa precisa de foto ou arte |

**Título da capa: uma frase só, com dois pesos** sobre o fundo petróleo:
- texto normal → **fino** (Poppins Light), branco
- `**palavra**` → **grosso** (Sora negrito), **âmbar**: é o destaque, a parte maior e mais forte

Exemplo: `"Manual de **exportação financeira**"`. Sem marcação, a skill faz sozinha: a parte antes de ` – `
(ou a 1ª palavra) fica fina e branca, e o resto vira destaque em âmbar. A parte fina precisa da fonte
**Poppins Light** instalada (Google Fonts).

**Espaço responsivo:** o título e a ficha andam juntos, com um espaço fixo entre eles, e o conjunto fica
centralizado na área petróleo. Nada de um vazio grande entre o subtítulo e a ficha.

**Pouco texto na capa:** título com até ~6 palavras, subtítulo curto (por exemplo, "Passo a passo · nov. 2026") e
ficha com no máximo 3 itens. O resto vai para dentro do documento. Sem filete embaixo do subtítulo.

**Capa feita fora (imagem):**
- Crie no tamanho **A4 em pé: 2480 × 3508 px** (300 dpi), em PNG ou JPG. Se a imagem vier em outra
  proporção, o script corta as sobras pelo centro, sem distorcer, então deixe o principal no meio.
- Use as cores da marca e a logo de `assets/`. O título pode já estar na arte, que é o caso padrão.
- Se a arte for só a foto, sem texto, use `"texto": True` para o script escrever o título, a logo e a
  ficha por cima, em branco. Nesse caso, use uma foto escura ou com área escura onde fica o texto.
- No Word, sem script: **Inserir → Imagem**, depois **Quebra de texto → Atrás do texto** e ajuste até as
  bordas da folha.

Trabalhos acadêmicos seguem a capa e a folha de rosto da ABNT, sem a faixa da marca. Fora isso, use a faixa ou a capa.

## 5. Seções (NBR 6024)

- **Numeração progressiva** em algarismos arábicos, com **um espaço** entre o número e o título e
  **sem ponto, hífen ou travessão**: `1 INTRODUÇÃO`, `1.1 Objetivo`, `1.1.1 Escopo`.
- Os títulos ficam **alinhados à esquerda** e separados do texto, antes e depois, por **uma linha em branco**.
- Títulos **sem número** ficam **centralizados**: REFERÊNCIAS, APÊNDICE, ANEXO, GLOSSÁRIO, SUMÁRIO.
- **Alíneas** (`a)`, `b)`, `c)`): recuadas, começam com minúscula e terminam com ponto e vírgula; a
  última termina com ponto. Use alíneas em texto formal (contrato, política). Em relatórios e
  comunicados, marcadores (•) também servem.
- Em trabalhos acadêmicos, cada seção primária começa em página nova. Em documentos da empresa, não é
  preciso.

## 6. Tabelas, gráficos e figuras (ABNT + visual das planilhas)

**Identificação em cima e fonte embaixo, sempre.** Fica assim:

> **Tabela 1 –** Custo mensal dos itens ativos
> *(tabela)*
> **Fonte:** planilha Custos de Tecnologia 2026.

- A designação é **Tabela** (dados numéricos), **Quadro** (informação em texto), **Gráfico** ou
  **Figura**, com numeração própria para cada tipo.
- A **fonte é obrigatória**, mesmo para dados próprios: "Fonte: Elaborado pela Saldus."
- Cite a tabela ou figura no texto antes dela: "A Tabela 1 mostra…".
- Visual da tabela, igual ao da planilha:
  - Cabeçalho com fundo petróleo, texto branco em Sora negrito e **filete âmbar** embaixo. Ele se
    repete quando a tabela passa de página.
  - Corpo em **zebra**, só com **linhas horizontais** finas e **sem bordas laterais**, o que também
    atende o padrão do IBGE que a ABNT segue.
  - **Total** com fundo `primaria-clara`, negrito e linha petróleo em cima.
  - Texto à esquerda, **números à direita**, datas e status centralizados.
  - Fonte 10, largura de 100% do texto, e linhas que não se partem entre páginas.
- **Gráfico:** largura total, sem moldura, sem grade e sem eixo de valores. O valor fica na ponta da
  barra, em formato BR. Use petróleo, e com várias séries a ordem é `#12424D`, `#FFB745`, `#3E8A96`,
  `#8CC2C9`, `#D9822B`, `#A7B4B7`. Sem degradê, sombra ou 3D.

## 7. Citações e referências

**Citações (NBR 10520:2023):**
- **Curta** (até 3 linhas): dentro do parágrafo, entre aspas duplas.
- **Longa** (mais de 3 linhas): parágrafo próprio, com recuo de 4 cm, fonte 10, entrelinha simples e
  sem aspas.
- **Chamada:** o sobrenome vai só com a inicial maiúscula (mudança de 2023): `(Silva, 2019, p. 12)`
  ou "Segundo Silva (2019, p. 12)…". Com quatro autores ou mais, pode usar *et al.*:
  `(Maciel et al., 2019)`. Escolha um jeito e mantenha no documento todo.
- Citação direta sempre leva o número da página.

**Referências (NBR 6023):** em seção própria, **REFERÊNCIAS** (centralizado, sem número), no fim do
documento. Ficam em ordem alfabética, alinhadas à esquerda, com entrelinha simples e uma linha em
branco entre elas. O sobrenome vai em MAIÚSCULAS e o título da obra em **negrito**:
`SILVA, João. **Título da obra**: subtítulo. São Paulo: Editora, 2019.`
Para sites, acrescente "Disponível em: <url>. Acesso em: 16 nov. 2026."

**Nunca invente uma referência, um número de lei ou uma citação.** Se não tiver a fonte real, peça ao
usuário ou deixe um marcador visível, como `[FONTE A CONFIRMAR]`.

## 8. Componentes da marca

Todos ocupam **100% da largura do texto**. Pense na página como um bloco único.

- **Cartões de resultado:** de 2 a 4 números principais, logo no começo. Têm a mesma largura e fundo
  `primaria-clara`. O rótulo vai em MAIÚSCULAS 8 pt, o valor em Sora 18 petróleo e o detalhe em 9 pt.
  **Só o número principal** fica em âmbar.
- **Caixa de destaque:** fundo `primaria-clara`, **barra âmbar à esquerda** e título opcional em Sora.
  Serve para premissas e conclusões. A variante *atenção* (fundo `entrada`, barra `alerta`) é para
  riscos e prazos. No máximo uma ou duas por página.
- **Ficha (rótulo/valor):** Cliente, Data, Versão, Responsável, em 2 ou 3 colunas. O rótulo vai em
  MAIÚSCULAS 8 pt cinza e o valor embaixo, em 11 pt.
- **Assinaturas:** "Cidade, 16 de novembro de 2026." alinhado à direita e, abaixo, uma linha de ~7,5 cm
  para cada pessoa, com o nome em negrito e o cargo em 10 pt cinza.

**Página de aviso (opcional):** uma página logo depois da capa. No topo fica uma **faixa âmbar de
alerta**, com um triângulo de alerta petróleo em cada ponta, "ATENÇÃO" grande em Sora e o tipo de aviso
embaixo (por exemplo, DOCUMENTO CONFIDENCIAL). Embaixo vem o **texto completo no padrão ABNT**, com
subtítulos em negrito petróleo (Uso restrito, É expressamente proibido, Responsabilidade) e
marcadores. Os modelos prontos são *confidencial*, *uso interno* e *minuta para revisão*, e também dá
para passar um texto próprio no mesmo formato. Sem capa, o aviso vira uma caixa de "atenção" curta logo
abaixo da faixa. Não prometa no aviso o que a empresa não faz (por exemplo, "rastreamento do
documento"), a menos que o usuário confirme que existe.

**Marca d'água (opcional):** o contorno do "S" da marca (`assets/simbolo-s-contorno.svg`) em bege bem
claro (`#F8EEDC`), grande, centralizado e vazando um pouco das bordas da folha. Fica atrás do texto em
todas as páginas, menos na capa. É só textura: não use outra cor nem deixe mais forte, porque o texto
precisa continuar fácil de ler.

**Manual: capturas de tela marcadas** (como nos manuais de passo a passo):
- Cada clique ganha um **número num círculo petróleo** com contorno branco e uma **seta petróleo**
  apontando para o item. O item clicado ganha um **contorno âmbar** arredondado.
- Se o número não couber sem cobrir a tela (por exemplo, um item do menu no topo), ele fica **fora do
  print**, numa margem branca acrescentada, com a seta entrando na imagem.
- **O mesmo número aparece no texto**, antes do nome do item: "Clique em ❶ **“Rural”** > ❷
  **“Orientações-Fomento”**". Use aspas e negrito no nome exato do botão. Separe o caminho com `>`.
- O print vai com **moldura fina cinza** e largura total. Pode receber "Figura N – …" e "Fonte: captura
  de tela de … (ano)" quando o manual seguir a ABNT à risca.
- Links clicáveis aparecem em petróleo sublinhado: "acesse [o portal](https://...)".
- Para marcar, olhe o print e estime as coordenadas em pixels do item (ou frações de 0 a 1). Depois
  gere e **confira a imagem marcada antes de colocar no documento**: a seta tem que tocar o item certo.

## 9. Números, datas e estrutura

- Moeda **R$ 1.234,56** · percentual **12,5%** · data **16/11/2026** em tabelas e "16 de novembro de 2026" no texto.
- Estrutura por tipo de documento:

| Tipo | Abertura | Estrutura sugerida |
|---|---|---|
| **Relatório** | faixa (capa se for longo) | 1 RESUMO (com cartões) → seções numeradas → próximos passos → REFERÊNCIAS |
| **Proposta comercial** | capa | Ficha → 1 CONTEXTO → 2 SOLUÇÃO → 3 ESCOPO (alíneas) → 4 INVESTIMENTO (tabela com total) → 5 PRAZOS → 6 CONDIÇÕES → assinaturas |
| **Contrato / termo** | título sóbrio, sem capa | Ficha das partes → cláusulas como seções ("1 OBJETO", "2 PRAZO"), com alíneas → assinaturas e testemunhas |
| **Ofício / carta** | faixa (título = assunto) | Local e data → destinatário → texto → assinatura |
| **Ata de reunião** | faixa | Ficha (data, horário, local, participantes) → 1 PAUTA → 2 DECISÕES (caixa) → 3 TAREFAS (tabela com responsável, prazo e status) |
| **Política / procedimento / norma** | título sóbrio com versão e vigência na linha de baixo (sem capa, sem ficha) | Seções → 1 OBJETIVO → 2 APLICAÇÃO → 3 REGRAS (alíneas) → histórico de versões (quadro) |
| **Comunicado** | faixa | Texto curto (até 1 página) e uma caixa de destaque com o que muda |
| **Manual / passo a passo** | capa (inteira ou imagem) ou faixa | 1 OBJETIVO → 2 ANTES DE COMEÇAR (acessos, requisitos) → uma seção por tarefa, cada passo com texto + **captura marcada** → caixa de dicas → 9 DÚVIDAS FREQUENTES |

## 10. Script Python (`scripts/estilo_word.py`)

Requer `pip install python-docx`. O `grafico_barras` também precisa de `matplotlib`.

```python
import sys; sys.path.insert(0, "<pasta da skill>/scripts")
from estilo_word import *

doc = novo_documento("Relatório de Custos", "Saldus · novembro de 2026")   # faixa + logo
# capa: novo_documento("Proposta Comercial", "Projeto X", capa={"info": {"Preparado para": "ACME",
#                      "Data": "16/11/2026", "Versão": "1.0"}})            # estilo bloco
#       capa={"estilo": "inteira", "info": {...}}                          # folha toda petróleo
#       capa={"estilo": "inteira", "foto": "foto.jpg", "info": {...}}      # + foto no alto
# título da capa: novo_documento("Manual de **exportação financeira**", ...)  # fino branco + destaque âmbar
#       capa={"imagem": "capa.png"}                                        # capa pronta feita fora
#       capa={"imagem": "foto.jpg", "texto": True, "info": {...}}          # foto + título por cima
titulo(doc, "Resumo")                  # -> "1 RESUMO" (numera sozinho); nivel=2,3,4 para subseções
texto(doc, "O custo caiu **29%**.")     # **negrito** em petróleo, _itálico_
cartoes(doc, [{"rotulo": "Custo mensal", "valor": "R$ 1.800,00", "detalhe": "6 itens", "destaque": True}])
tabela(doc, ["Item", "Custo mensal", "Status"], linhas, formatos={"Custo mensal": "moeda"},
       coluna_status="Status", total=True, titulo_tabela="Custo mensal por item",
       fonte="planilha Custos de Tecnologia 2026.")       # designacao="Quadro" para texto
grafico_barras(doc, categorias, valores, "Custo mensal por item", fonte="Elaborado pela Saldus.")
figura(doc, "foto.png", "Fachada da sede", fonte="Acervo Saldus.")
citacao(doc, "Trecho com mais de três linhas...", "(Silva, 2019, p. 12)")
alineas(doc, ["primeiro item", "segundo item"])          # a) ...; b) ....
lista(doc, ["item solto", "outro"])                       # marcadores
destaque(doc, "Texto.", titulo="Premissa")                # tipo="atencao"
ficha(doc, {"Cliente": "ACME", "Data": "16/11/2026"}, colunas=3)
assinaturas(doc, [("Nome", "Cargo · Saldus")], local_data=f"Cidade, {data_extenso(hoje)}.")
referencias(doc, ["SILVA, João. **Título**. São Paulo: Editora, 2019."])
salvar(doc, "relatorio.docx")

# Documento formal: novo_documento("Política de X", "Assunto · Versão 1.0 · Vigência ... · Confidencial")
# Padrão digital: sem cabeçalho + rodapé em faixa (logo + número). Variações, só se pedirem:
#   novo_documento(..., estilo_rodape="linha")   rodapé simples · paginacao=False sem número
#   novo_documento(..., cabecalho=True)          cabeçalho no alto (logo + número + filete)
# Marca d'água: novo_documento(..., marca_dagua=True)            # "S" em contorno no fundo das páginas
# Aviso: novo_documento(..., capa={...}, aviso="confidencial")   # "uso_interno", "minuta" ou texto próprio
#   texto próprio: aviso={"titulo": "Atenção", "subtitulo": "Documento confidencial",
#                         "blocos": [("p", "parágrafo"), ("h", "Subtítulo"), ("li", ["item;", "item."])]}
# Manual: texto com marcadores {n} e links; captura com as marcações
texto(doc, "Acesse [o portal](https://portal.com.br) e clique em {1} **“Rural”** > {2} **“Orientações-Fomento”**.")
captura(doc, "print.png", [
    {"n": 1, "caixa": (330, 10, 410, 40), "lado": "cima"},       # contorno âmbar + número acima
    {"n": 2, "caixa": (330, 455, 485, 490), "lado": "esquerda"},  # número à esquerda, seta →
], titulo_figura="Menu Rural do portal", fonte="captura de tela do portal (2026).")
# só a imagem marcada (para conferir ou usar em outro lugar): marcar_imagem("print.png", marcas)
```

- Formatos de coluna: `moeda`, `inteiro`, `decimal`, `percentual` (0,125 → 12,5%), `variacao`, `data`,
  `mes`, `centro`. Passe **números de verdade**. Eles saem formatados, alinhados à direita e somados
  no total.
- `titulo(doc, "Referências", numerar=False)` cria um título centralizado sem número.
- `python scripts/estilo_word.py exemplo.docx` gera um exemplo.
- **Modelos prontos** em `assets/modelos/`, para abrir no Word e preencher. Quando o pedido for de um
  desses tipos, comece do modelo, ou da função `modelo(tipo, caminho)`, e troque os textos de exemplo:

| Modelo | O que já vem pronto |
|---|---|
| `modelo-documento.docx` | **Documento formatado**: capa inteira com título fino + destaque âmbar, cabeçalho completo, resumo com cartões, tabela com total, gráfico, subseção, conclusão em caixa e referências |
| `modelo-documento-simples.docx` | Documento curto: faixa de título, seções, tabela, caixa e referências |
| `modelo-manual.docx` | **Manual**: capa inteira, aviso confidencial, marca d'água, sem cabeçalho e com rodapé em faixa, objetivo, pré-requisitos, passo com print marcado (troque pelo seu print), dica e dúvidas frequentes |
| `modelo-formal-impressao.docx` | **Documento formal** (política, contrato, termo): faixa de título, versão e vigência na linha abaixo do título, cláusulas numeradas, rodapé em faixa e assinatura com campos em branco |
| `modelo-proposta.docx` | Proposta: capa em bloco com ficha, contexto, solução, escopo em alíneas, investimento com total, condições e assinaturas |

  Para gerar de novo: `python scripts/estilo_word.py --modelo manual saida.docx` (ou `documento`,
  `documento-simples`, `proposta`).

## 11. Conferir antes de entregar

- **Sempre olhe o resultado** antes de entregar: converta para PDF (Word ou `soffice --headless
  --convert-to pdf`) e veja as páginas como imagem. Confira capa, quebras de página, tabelas, prints
  marcados e rodapé.
- Nenhuma página em branco a mais: a capa calcula os espaços para caber numa folha só. Se mudar textos
  da capa, confira de novo.
- O LibreOffice não usa Sora/Poppins e desenha emendas finas entre células petróleo. Isso não aparece no
  Word, que é a referência final.
- Imagens dentro do texto (logo, ícones) nunca vão em linha com altura **exata**, porque o Word corta a
  parte de baixo. Use altura **mínima**.
- Fotos e capas são cortadas pelo centro, sem esticar. Avise o usuário se o assunto da foto estiver
  fora do centro.

## 12. Google Docs

Use `scripts/estilo_docs.gs`: abra **Extensões → Apps Script**, cole o arquivo, salve e recarregue.
Vai aparecer o menu **Estilo**, com *Aplicar estilo padrão* (margens Normal, fonte 12, entrelinha 1,5,
recuo, justificado, títulos e tabelas), *Inserir faixa de título* e *Inserir caixa de destaque*. O
número de página entra por **Inserir → Números de página** (canto superior direito), e a logo por
**Inserir → Imagem**. No Google Docs só existe o estilo básico. Capa inteira, foto, aviso, marca
d'água, rodapé em faixa e prints marcados são feitos pelo script Python, gerando um `.docx` que pode ser
aberto no Google Docs.

## 13. Checklist final

- [ ] Margens Normal do Word (2,5 cm em cima/embaixo, 3 cm nas laterais), fonte 12, entrelinha 1,5, recuo de 1,25 cm, texto justificado
- [ ] Seções numeradas sem ponto (`1 INTRODUÇÃO`, `1.1 Objetivo`), alinhadas à esquerda; sem número = centralizado
- [ ] Sem cabeçalho; rodapé em faixa com logo e número da página (oculto na capa)
- [ ] Documento formal: sem capa
- [ ] Capa: pouco texto, título com parte fina branca + destaque âmbar, sem vazio grande, sem página extra, logo inteira
- [ ] Tabelas, gráficos e figuras com "Tabela N – título" em cima e "Fonte:" embaixo, citados no texto
- [ ] Citações longas com recuo de 4 cm, fonte 10, simples, sem aspas; chamadas como (Silva, 2019, p. 12)
- [ ] REFERÊNCIAS no fim, em ordem alfabética, sem nenhuma inventada
- [ ] Faixa petróleo (ou capa) com logo; âmbar só em filetes, barras e no máximo um cartão
- [ ] Tudo com 100% da largura do texto; números e datas no formato brasileiro
- [ ] Manual: cada passo com o número no texto **e** no print, e as setas apontando para o item certo
