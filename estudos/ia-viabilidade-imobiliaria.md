# IA aplicada à viabilidade de loteamentos

**Guia de estudo e laboratório prático**
**Versão:** 1.0 — agosto de 2026
**Para:** quem já domina VPL, TIR, múltiplo e payback, e quer aprender a usar IA nesse trabalho.

> **Aviso.** Este documento é material de estudo. Não é recomendação de investimento, não é
> parecer contábil nem jurídico. Os números do caso usado no laboratório são fictícios e servem
> só para ensinar o método. Tributação, estrutura societária e enquadramento do empreendimento
> precisam ser confirmados com contador e advogado antes de qualquer decisão.

---

## Sumário

- [Parte 0 — Por que "qual IA calcula melhor" é a pergunta errada](#parte-0)
- [Parte 1 — O método: os quatro papéis da IA](#parte-1)
- [Parte 2 — Excel ou Google Sheets](#parte-2)
- [Parte 3 — Engenharia de prompt para trabalho numérico](#parte-3)
- [Parte 4 — O modelo de loteamento](#parte-4)
- [Parte 5 — Laboratório: seis exercícios](#parte-5)
- [Parte 6 — Trilha de aprofundamento](#parte-6)
- [Parte 7 — Régua de confiança](#parte-7)

---

<a name="parte-0"></a>
## Parte 0 — Por que "qual IA calcula melhor" é a pergunta errada

Você perguntou qual IA funciona melhor com números. A resposta honesta é desconfortável:
**nenhuma é boa o suficiente para ser sua calculadora**, e a diferença entre elas é menor do
que a diferença entre um método bom e um método ruim.

Os dados de 2026 são públicos. O *FinSheet-Bench* mediu modelos de ponta respondendo perguntas
numéricas sobre planilhas financeiras reais de fundos. O melhor resultado foi de **82,4%** de
acerto. Os quatro primeiros colocados ficaram empilhados entre 80,2% e 82,4% — ou seja,
empatados dentro do ruído. E o dado que importa de verdade: **na planilha maior do conjunto, a
média de acerto de todos os modelos caiu para 48,6%**.

Traduzindo para a sua realidade: numa planilha pequena, a IA erra cerca de uma resposta em
cinco. Numa planilha do tamanho de um estudo de viabilidade de loteamento — 180 meses de
fluxo, dezenas de linhas, três abas — ela erra **perto de metade**. Uma moeda.

Agora perceba a armadilha. Ela não erra de um jeito que pareça erro. Ela devolve
`TIR de 27,3% a.a.`, com uma casa decimal e ar de convicção. Um número redondo demais, plausível
demais, e sem nenhum rastro de como foi obtido. É exatamente esse tipo de número que entra numa
reunião e sai como decisão de comprar uma gleba de nove milhões.

### O que muda quando você aceita isso

Se a IA não é a calculadora, ela é o quê? Ela é **as outras quatro coisas** que um analista faz
e que ninguém percebe que são o trabalho de verdade: desenhar a estrutura do modelo, escrever
as fórmulas, auditar o que está pronto, e brigar com as premissas. A conta em si — a soma, o
desconto do fluxo, a TIR — é a parte trivial, e é a única que um motor determinístico faz com
100% de acerto desde 1985.

O erro que quase todo mundo comete é usar a IA justamente na parte trivial, onde ela é fraca, e
não usar nas quatro partes onde ela é forte.

### Então, o comparativo prático que você pediu

Reenquadrado. Não "qual é a melhor", mas **qual serve melhor para cada papel**:

| Papel | O que importa no modelo | Comentário |
|---|---|---|
| **Arquiteta** do modelo | Raciocínio longo, memória do contexto, capacidade de discutir e revisar | Qualquer modelo de topo com modo de raciocínio estendido resolve. Aqui a diferença entre eles é pequena. |
| **Escritora de fórmulas** | Conhecimento das funções da planilha, precisão sintática | Todos os modelos de topo acertam `XLOOKUP`, `XTIR`, `XVPL`, `SOMASES`. Erram em fórmulas matriciais complexas — desconfie delas. |
| **Auditora** de planilha pronta | **Ler o arquivo inteiro**, não um recorte colado | Aqui a diferença é grande, e não é o modelo: é a **ferramenta**. Vale mais um modelo mediano que abre seu `.xlsx` de verdade do que um modelo excelente lendo texto colado. |
| **Calculadora** | — | ❌ Nenhum. Sem exceção. |
| **Crítica / red team** | Disposição de discordar em vez de concordar | O ponto fraco de todos: são treinados para ser agradáveis. Só funciona com prompt que **obriga** a discordar. Ver Parte 3. |

Três observações que valem mais que o ranking:

1. **Raciocínio estendido ligado, sempre.** O modo "thinking"/"reasoning" é a diferença entre um
   modelo que acerta 80% e o mesmo modelo acertando bem menos. Não é opcional em trabalho numérico.
2. **Modelo que executa código vence modelo que "pensa" a conta.** Quando a IA escreve um trecho
   de código, roda e te devolve o resultado, a aritmética passou por um interpretador — e aí ela
   é exata. Quando ela responde de cabeça, é estatística. A mesma marca, no mesmo dia, tem os
   dois comportamentos: depende de você pedir.
3. **Isso muda a cada trimestre.** Não decore um ranking. Decore o teste: dê a mesma pergunta
   difícil sobre a sua planilha real a dois modelos e confira contra a célula. Em quinze minutos
   você sabe mais sobre qual serve para o seu caso do que qualquer benchmark.

---

<a name="parte-1"></a>
## Parte 1 — O método: os quatro papéis da IA

Esta é a parte central do documento. Se você ler só uma seção, leia esta.

### 1. Arquiteta — antes de existir número

O uso de maior retorno e o menos praticado. Antes de abrir a planilha, você conversa com a IA
sobre **qual deve ser a estrutura do modelo**: quais linhas o fluxo precisa ter, quais premissas
são independentes e quais são derivadas, onde estão os descasamentos de prazo, qual métrica
decide.

Isso funciona bem porque é raciocínio verbal sobre estrutura — o ponto forte do modelo — e porque
não há nenhum número para ela errar. Você sai com o esqueleto do modelo antes de digitar a
primeira célula, e o esqueleto é onde moram os erros caros. Uma linha de fluxo esquecida custa
mais do que mil erros de arredondamento.

> Erro de arquitetura você descobre em seis meses. Erro de fórmula você descobre na conferência.

### 2. Escritora de fórmulas — a tradução

Você descreve a regra de negócio em português; ela devolve a fórmula da célula. *"A parcela do
mês só entra se o lote já foi vendido, se ainda está dentro das 120 parcelas, e se o comprador
não distratou."* Ela escreve o `SE` aninhado ou o `SOMASES` correspondente.

O ganho não é velocidade de digitação. É que **você revisa uma fórmula muito mais rápido do que
escreve uma**, e revisar é onde seu conhecimento de negócio entra. A IA sabe sintaxe; você sabe
que a parcela é corrigida por INCC durante a obra e por IPCA depois. A divisão é natural.

### 3. Auditora — o papel de maior ROI

Este é o papel mais subestimado, e é onde eu colocaria metade do seu tempo de aprendizado.

Você tem uma planilha pronta — sua, ou de um corretor, ou de um sócio, ou de um parceiro
querendo vender a gleba. Você entrega o arquivo para a IA e pede que ela **procure o erro**.
Não que ela recalcule: que ela procure inconsistência de sinal, referência que não acompanhou a
linha, premissa que aparece com dois valores diferentes em duas abas, período que ficou de fora
do somatório, taxa anual multiplicada por doze onde deveria ser elevada a doze.

Por que ela é boa nisso: encontrar padrão quebrado é reconhecimento, não aritmética. É
exatamente o que um modelo de linguagem faz bem. E é chato o bastante para que nenhum humano
faça direito na quinta revisão.

Um detalhe operacional que muda tudo: a IA precisa ver **as fórmulas**, não os resultados. Uma
planilha aberta como arquivo mostra as duas coisas. Um print de tela mostra só o resultado — e
aí ela não pode auditar nada, só concordar com você.

### 4. Intérprete e crítica — o sócio que discorda

Modelo pronto, números na tela. Agora a IA lê o resultado e ataca: *"A velocidade de 12 lotes por
mês sustentada por 58 meses seguidos é o dobro da média histórica da região? Se cair para 8, o
que acontece com a exposição máxima?"*

Este papel só funciona se você **obrigar** a IA a discordar, porque o comportamento padrão dela
é concordar com você. Prompt para isso na Parte 3.

### 5. Calculadora — nunca

E agora a regra que amarra os quatro papéis, que é a única coisa deste documento que eu pediria
para você decorar:

> ### Toda conta vive numa célula ou em código executado.
> ### Nenhum número que a IA "disse" entra numa decisão sem estar amarrado a uma célula que você consegue clicar e ver a fórmula.

Se a IA te der um VPL no meio de um parágrafo de chat, esse número não existe. Ele não é
auditável, não é reprodutível, não sobrevive a uma pergunta do seu contador, e tem chance real
de estar errado. Peça a fórmula, coloque na planilha, e leia o resultado da planilha.

Isso não é desconfiança exagerada. É a mesma disciplina que você já aplica quando não aceita o
"mais ou menos isso" de um corretor sem ver a planilha dele.

---

<a name="parte-2"></a>
## Parte 2 — Excel ou Google Sheets

Você pediu para decidir, e é uma pergunta real. Não há resposta universal — a resposta é
**qual papel você está jogando agora**. O melhor arranjo é usar os dois, cada um no seu lugar.

### Excel como motor

É aqui que você constrói o modelo de verdade. Por quê:

1. **Tabela bidimensional visual.** Um loteamento com 180 meses de fluxo e 50 linhas é uma
   matriz 180 × 50 que você consegue rolar e conferir com os olhos. Em Sheets, com abas de
   seleção interativa, fica opaco — perde você metade do poder que a visualidade dá.

2. **Solver, Atingir Meta, Tabelas de Sensibilidade.** Tudo aqui é nativo, rápido e exato.
   Sheets é fraco nisso. Você consegue fazer, mas é mais lento e menos legível.

3. **A IA consegue auditar o arquivo inteiro.** Salve seu `.xlsx`, dê para Claude ou ChatGPT com
   code interpreter, e ele abre o arquivo, lê todas as abas, todas as fórmulas, e critca de
   verdade. A fórmula `=SUMIF(...)` que você deixou referenciando a aba errada? Ele acha.

4. **Sem conexão de internet, você trabalha.** Se você está numa viagem, ou no canteiro, ou a
   rede caiu, você abre o arquivo no notebook e segue.

**Como trabalhar:** edição direta; números locais; backup manual ou sincronização com OneDrive.

### Google Sheets como camada de distribuição

Não para editar o modelo de novo. Para mostrar, comentar, coletar feedback de sócios e do
contador. Por quê:

1. **Colaboração real.** Sócio abre o link, vê os números, comenta: *"A velocidade aqui está
   otimista"*. O comentário fica amarrado à célula. Você revisa, volta ao Excel, muda,
   reexporta. Feedback estruturado.

2. **Versionamento.** Sheets guarda histórico automático. Você não vai perder uma versão antiga
   se sua máquina explodir.

3. **Móvel, browser, sem software.** Se alguém precisa conferir um número num celular, consegue.

**Como trabalhar:** copie a tabela principal do Excel para uma aba de Sheets (data + tabelas). O
resto (sensibilidade, Monte Carlo, crítica) fica no Excel. Sheets é só a verdade principal,
resumida, bonita.

### Fluxo de ida e volta que funciona

```
1. Excel: monta e calcula tudo
2. Excel: salva .xlsx
3. Você: vai para Sheets, limpa tudo, copia só a tabela principal
4. Você: compartilha o link de Sheets com sócio/contador
5. Contador comenta: "essa premissa está errada"
6. Você: volta ao Excel, corrige a premissa
7. Excel: regenera tudo (porque as fórmulas estão lá)
8. Você: copia os novos números para Sheets
9. Sheets: contador vê a cor mudar, é notificado, relê
```

**O que NÃO fazer:**

- Não monte o modelo em Sheets e edite direto nele. Vai ser fraco, vai demorar, vai perder auditoria.
- Não copie/cole um intervalo de Sheets inteiro no chat e peça para a IA ler. Ela vai errar porque
  enxerga só texto, não estrutura. É o pior dos mundos: lenta para ler, fácil de errar.
- Não use Sheets como "guarda-chuva" para quatro Excels diferentes de quatro sócios. Sincroniza
  com o tempo, dá conflito, e ninguém sabe qual é a versão certa.

### Resumo da divisão

| Tarefa | Ferramenta | Por quê |
|---|---|---|
| Desenhar o modelo | Chat + Excel | O chat ajuda na arquitetura; o Excel é onde fica |
| Montar as fórmulas | Excel | Direto; a IA faz a fórmula, você cola e revisa |
| Auditar pelo IA | Excel + chat (arquivo) | A IA precisa do arquivo para auditar de verdade |
| Rodar sensibilidade | Excel | Solver, tabelas bidimensionais |
| Mostrar para terceiro | Sheets | Colaboração, comentário, móvel |
| Receber feedback | Sheets (comentários) | Comentário estruturado, rastreável |
| Incorporar feedback | Excel | Volta ao motor e regenera |

---

<a name="parte-3"></a>
## Parte 3 — Engenharia de prompt para trabalho numérico

Aqui estão as seis técnicas que realmente fazem diferença. Cada uma está escrita com o
**antes/depois** do prompt, para você ver exatamente o que muda.

A primeira coisa a entender é que "engenharia de prompt" é só a aplicação da Parte 1 — os quatro
papéis — traduzida em instruções bem-formadas. Não é magia de linguagem. É precisão.

### Técnica 1: Contrato de premissas

**O problema que resolve:** a IA inventa número quando não tem. Você pede TIR de um loteamento,
ela não tem a velocidade de vendas, e devolve "um valor razoável" — 8 lotes por mês. Esse número
entra na planilha e ninguém percebe que veio de coisa nenhuma.

**O antes:**

```
Preciso da TIR mensal de um loteamento em Palmas com 50 lotes, 
preço médio R$ 200 mil, gleba comprada à vista por R$ 5 milhões.
Qual é?
```

Resposta: *TIR de 3,2% a.m. (equivalente a 46% a.a.)*

Você não sabe se ela assumiu 6 lotes por mês ou 20. A premissa é invisível.

**O depois:**

```
Vou te dar um loteamento para simular. Não invente número nenhum.
Se não tiver uma premissa, escreva [FALTA: descrição] e vou providenciar.

As premissas são:
- Nº de lotes: 50
- Preço médio por lote: R$ 200 mil
- Custo de aquisição: R$ 5 milhões (à vista)
- Velocidade de vendas: [FALTA: lotes/mês?]
- Custo de infra total: [FALTA: R$?]
- Cronograma de desembolso de infra: [FALTA: qual?]
- Entrada na venda: [FALTA: %?]
- Parcelas de venda: [FALTA: número e prazo?]

Antes de calcular a TIR, confirme que todas as premissas estão preenchidas.
```

Resultado: a IA sinaliza o que falta, em vez de fazer de conta que sabe. Você preencheu, ela
recalcula com números reais.

**Como aplicar isso:** sempre que começar um novo modelo, escreva o "contrato". Pegue a
IA ajudando — ela é boa em lembrar de coisas. Depois, congele a lista e reenvia sempre que
mexer no modelo.

### Técnica 2: Separação premissa / fórmula / resultado em abas distintas

**O problema que resolve:** auditoria fica impossível se premissa e resultado estão misturados
com fórmula. A IA (ou um humano) relê e não consegue distinguir o que foi assumido do que foi
calculado.

**O antes:** tudo na mesma aba, linhas soltas.

```
Aquisição: R$ 5.000.000
Nº de lotes: 50
Preço/lote: R$ 200 mil
VGV: =E3 * E4
Custo infra/lote: R$ 100 mil
Infra total: =E5 * E4
...
TIR: [uma fórmula XTIR obscura que não se vê ]
```

Quando a IA tenta auditar, ela se perde em qual linha é premissa e qual é resultado.

**O depois:** três abas, hierarquicamente:

```
Aba 1 — PREMISSAS
  Aquisição do terreno: R$ 5.000.000
  Nº de lotes: 50
  Preço/lote (médio): R$ 200.000
  Custo infra/lote: R$ 100.000
  Entrada de venda: 15%
  Nº de parcelas: 120
  Prazo parcelas (meses): 360
  Correção juros (a.m.): 0,70%
  Velocidade venda (lotes/mês): 1,0

Aba 2 — CALCULADOS
  VGV: =PREMISSAS!B2 * PREMISSAS!B3 (com referência explícita)
  Infra total: =PREMISSAS!B6 * PREMISSAS!B3
  Caixa inicial: -PREMISSAS!B2
  ...

Aba 3 — RESULTADO
  TIR mensal: [lê só Aba 2, sem fórmulas complexas locais]
  VPL: [idem]
  Exposição máxima: [idem]
  Payback meses: [idem]
```

Por que funciona: a IA (e você) conseguem auditar em três passes. Aba 1 é uma lista de premissas
— óbvio ler e criticar. Aba 2 mostra as derivações — fácil achar erro de lógica. Aba 3 mostra
decisão — você consegue rolar a página inteira de RESULTADO em 30 segundos.

### Técnica 3: "Mostre a fórmula, não o resultado"

**O problema que resolve:** a IA te dá a resposta (TIR = 31%), você não consegue auditar como
chegou lá, e se estiver errada, você não vê por quê.

**O antes:**

```
Qual é a TIR mensal dessa série de fluxos?
[você cola uma coluna com 180 números]

Resposta: TIR mensal: 2,47% (equivalente a 34,2% a.a.)
```

Você não sabe se ela usou a fórmula certa, se contou os períodos bem, ou se arredondou sem
contar.

**O depois:**

```
Calcule a TIR mensal dessa série de fluxo. 
Antes de me dar o resultado, mostre:
1. A fórmula que você usou (ex.: =XTIR(intervalo, período_inicial, período_final))
2. O intervalo exato (ex.: FLUXO!B5:B184)
3. O resultado

Se não conseguir montar a fórmula com certeza, diga qual é a dúvida.
```

Resultado: você vê a fórmula, consegue bater contra seu Excel, e se estiver errada, acha na
hora.

### Técnica 4: Verificação cruzada

**O problema que resolve:** um número certo por um caminho errado. Exemplo: VGV de 50 × 200 mil
= 10 milhões, certo. Mas e se a IA tiver contado 51 lotes sem avisar, errado.

**O antes:**

```
VGV = 50 lotes × R$ 200 mil = R$ 10 milhões.
TIR = ... [ela calcula]
```

**O depois:**

```
VGV = 50 lotes × R$ 200 mil.
Confirme: qual é o valor total?
Agora calcule de novo: R$ 10.000.000 / R$ 200 mil = quantos lotes?
Se o resultado for 50, seguimos. Se for diferente, paramos e achamos o erro.

Depois recalcule a TIR usando o VGV confirmado.
```

É mecânico, é um pouco chato, mas pega metade dos erros de arredondamento e lógica antes de
virar verdade na planilha.

### Técnica 5: Teste de amarração

**O problema que resolve:** a planilha parece coerente, mas está quebrada (fluxo de entrada
diferente de fluxo de saída, caixa inicial não bate caixa final, etc.).

**O antes:** você rola a planilha de ponta a ponta e tenta não errar.

**O depois:** uma linha — ou melhor, uma célula — que é seu "teste de amarração":

```
Aba VALIDAÇÃO:
  Caixa inicial: [referência a RESULTADO]
  + Entradas de vendas: [somatório de FLUXO da seção de receita]
  - Saídas (infra + outros): [somatório de FLUXO da seção de custo]
  = Caixa final calculada: [fórmula]
  
Caixa final (do fluxo mês 180): [referência direto ao FLUXO]

Validação: Caixa final calculada = Caixa final do fluxo?
  (se SIM, célula fica verde; se NÃO, fica vermelha)
```

Por que é poderoso: quando você muda uma premissa — digamos, o número de lotes — a célula de
validação fica vermelha na hora, avisando que o modelo quebrou em algum lugar. Você acha o
erro mais rápido.

### Técnica 6: Red team de premissas

**O problema que resolve:** você está otimista, mas não sabe onde. A IA concorda com você por
padrão.

**O prompt que força a discordar:**

```
Você é um sócio cético. Vou te dar um estudo de viabilidade de loteamento. 
Seu trabalho é ASSUMIR que o estudo está ERRADO e PROCURAR as premissas frágeis.

Para cada seção (aquisição, infra, velocidade de vendas, etc.), responda:
1. Qual seria a pior situação razoável?
2. Essa premissa tem história, ou é apalpação?
3. Se cair 20% dessa premissa, a TIR muda pra quanto?

Não seja educado. Seja crítico. Aponte os três maiores riscos.
```

O resultado não é a verdade — é um teste de vulnerabilidade da sua tese. Se a TIR é 35% e cai
para 15% com queda de 20% na velocidade de vendas, você sabe que velocidade é o risco crítico.

### Anti-padrões que você provavelmente está cometendo

Aqui estão as coisas que quase todo mundo faz e que destroem a qualidade do resultado:

1. **Pedir TIR de cabeça, sem arquivo.** *"Tenho um loteamento de 50 lotes, preço médio R$ 200
   mil, infra de R$ 5 milhões, velocidade 2 lotes/mês. Qual é a TIR?"* Resposta: adivinhação. A IA
   está inventando o fluxo dela e devolvendo um número que não é auditável. Sempre: arquivo.

2. **Tabela grande colada em texto.**
   ```
   Mês | Vendas | Custo | Saldo
   1   | 200    | 500   | -300
   2   | 400    | 500   | ...
   ```
   Resultado: a IA lê como tabela desestruturada, erra linha, erra somatório. Sempre arquivo.

3. **Aceitar número sem célula.** A IA: *"Margem sobre VGV: 32%"*. Você: ótimo. Errado. Pergunta:
   *"Mostre a fórmula."* Se não há fórmula, o número não existe.

4. **Prompt sem unidade nem período-base.** *"Qual é o custo de infra?"* — custo em R$ ou em R$/
   m²? Por quantos meses? Sem responder essas duas perguntas, a IA chuta. Sempre: unidade, período,
   referência de premissa.

5. **Editar a planilha sem comunicar à IA.** Você mudou a velocidade de vendas no Excel, mas não
   aviou a IA. Ela segue calculando com o número antigo, em um outro chat. Dois modelos diferentes
   andando em paralelo. Sincronize: sempre que mudar uma premissa de verdade, repasse para a IA
   antes de pedir cálculo novo.

6. **Confiar na IA para tributação.** *"Com LSJP, qual é o imposto sobre o lucro?"* Resposta: a
   IA vai adivinhar. Resultado: você entra numa reunião com contador com a premissa errada.
   Sempre: comente, sempre, *"isso precisa ser confirmado com contador"* e deixe uma célula
   vazia ou marcada com [PENDÊNCIA TRIBUTÁRIA].

---

<a name="parte-4"></a>
## Parte 4 — O modelo de loteamento (a substância do domínio)

Aqui está a anatomia de um modelo de loteamento que resiste. Não é um roteiro de seis passos
("abra uma planilha, coloque premissas…"). É a estrutura conceitual que faz o modelo funcionar.

Se você está vindo de incorporação vertical, alerta: **loteamento não é incorporação com outro
nome**. O ritmo de caixa é diferente, o risco principal é diferente, e a série de decisões que
determinam viabilidade é diferente. Incorporação é sobre margem. Loteamento é sobre
descasamento.

### A tese central: o descasamento de caixa

Em um loteamento bem-sucedido, tudo sai à vista (infra), e entra parcelado (venda de lotes).
Você paga terraplenagem, drenagem, pavimentação, iluminação, água, esgoto — dezenas de milhões
— em oito a doze meses. A receita entra assim: R$ 30 mil de entrada por lote (15% de R$ 200 mil)
no mês um, e mais R$ 14 mil por mês pelos 120 meses seguintes (o que sobra, parcelado em
prestações mensais sem defasagem).

O resultado é que a gleba começa negativa, monton nos meses 2–10 (infra saindo, receita sendo
parcelada), e só fica positiva lá pelo mês 60 se tudo der certo.

A métrica que importa **não é TIR**. É **exposição máxima de caixa** — no pior mês, quanto você
precisa ter em caixa para não quebrar. Se você tem R$ 3 milhões e a exposição é R$ 4 milhões,
o projeto não sai. É matemático.

Uma incorporação pode ter TIR de 35% e estar viva em caixa (porque a venda sai escalonada e a
obra desce escalonada). Um loteamento pode ter TIR de 35% e estar morto em caixa (porque toda
infra sai na frente). Já que isso não é óbvio, a primeira coisa que você coloca na planilha é:

```
Aba RESULTADO:
  Exposição máxima de caixa: [mín do fluxo acumulado]
  Mês da exposição: [período em que ocorre]
  Caixa em poder: [sua entrada de capital, ou linha de crédito]
  Status: [= SE(exposição < caixa, "viável", "quebrado")]
```

Se status = "quebrado", a TIR é irrelevante. Você parou aqui.

### Aquisição do terreno

Nem toda gleba é comprada à vista. Você tem três situações distintas:

**Compra à vista:**
- Saída no mês zero: valor total
- Mais simples; não há financiamento de terreno

**Compra a prazo financiada:**
- Saída parcelada: entrada + parcelas (digamos, 30% à vista, 70% em 24 meses)
- Melhora exposição de caixa
- Juros e correção: você paga IPCA + juros ou uma taxa fechada?
- Linha no fluxo: "Desembolso financiamento terreno"

**Permuta física:**
- Você não compra; você vira parceiro com o dono do terreno
- Ele recebe uma quantidade de lotes como pagamento (digamos, 8 dos 50 lotes)
- Implicação: a receita de vendas cai (só vendem 42 lotes), mas a saída de caixa de aquisição cai também
- Distorção estatística: a TIR sobe (menos capital investido), mas pode ser ilusória — você não
  tem receita nesses 8 lotes de verdade
- Como modelar: uma aba à parte com "Cenário sem permuta" e "Cenário com permuta", lado a lado,
  de forma que a diferença entre elas é só a quantidade de lotes vendidos e as receitas associadas

Permuta é um assunto em si e merecia capítulo separado. O ponto curto: **sempre modele permuta
como um cenário, não como cenário base**. Porque desconta lotes de receita, as métricas mudam
de forma não-óbvia.

### Aprovação, licenciamento, seus prazos e custos

O projeto não sai no mês zero da viabilidade. Antes de começar a pagar infra, você precisa:

1. Aprovação da Prefeitura (Lei 6.766, Lei Complementar municipal): 4–24 meses conforme município
2. Licença ambiental (quando necessária): 3–18 meses
3. Registro no Cartório: 1–3 meses **depois da obra estar pronta** (é uma saída de caixa no final!)
4. CCIR, SNCR, inscrição na prefeitura: mês a mês

**Na sua planilha, isso é:**

```
Aba CRONOGRAMA:
  Mês 1–6: Aprovação (zero saída de caixa, só custos legais)
  Mês 7–18: Infra (saída pesada)
  Mês 19+: Venda (entrada parcelada)
  Mês final+1: Registro em cartório (saída pequena)

Aba FLUXO:
  Linhas específicas: custos de aprovação, custos de registro
  Períodos: alinhados ao cronograma, não alinhados às vendas
```

Por que importa: você não pode começar a vender antes da aprovação. O prazode aprovação é fixo e
não reage à sua vontade. Planejamento ruim aqui viabilidade inteira.

### Infraestrutura: a curva física vs. curva financeira

A infraestrutura de um loteamento tem componentes que saem em momentos diferentes:

- **Terraplenagem, drenagem:** mês 1–3 (mobilização e alteamento)
- **Pavimentação:** mês 4–10 (depois que o terreno está nivelado)
- **Redes (água, esgoto, elétrica):** mês 3–12 (paralela à pavimentação)
- **Iluminação:** mês 11–12
- **Acabamento (ajustes, limpeza):** mês 13

**Na planilha:**

```
Custo de infra por lote: [premissa no PREMISSAS]
Custo total de infra: Custo/lote × Nº de lotes

Distribuição do custo no tempo:
  Mês 1–3: 15% (terraplenagem)
  Mês 4–10: 50% (pavimentação + redes)
  Mês 11–13: 20% (redes finais + iluminação)
  Mês 14: 15% (ajustes)
```

Essa é a "curva física" — como o trabalho acontece no chão. Não é linear. A maioria das
planilhas caseiras assume linear (10% ao mês durante 10 meses), e aí a IA descaseia com a
realidade de cronograma.

Agora, a "curva financeira" — quando você paga. Pode ser igual à curva física (você paga
conforme a obra anda), ou pode ser adiantada (você paga à vista e o empreiteiro controla o
ritmo). Escolha depende do seu poder de barganha. Modele as duas, lado a lado.

```
Cenário 1: Pagamento conforme a obra (curva física)
Cenário 2: Pagamento adiantado (tudo nos meses 1–2)
```

Custo-benefício: pagamento adiantado piora sua caixa, mas é barganha para desconto.

### Receita: lotes vendidos, entrada, parcelas

A receita é a série mais importante e a mais sujeita a erro.

**Premissas que alimentam:**

1. **Preço por metro quadrado do lote:** [R$/m²] (não: R$ por lote, porque o tamanho do lote pode
   variar)
2. **Área média do lote:** [m²]
3. **Preço médio por lote:** = Premissa 1 × Premissa 2
4. **Velocidade de vendas:** [lotes/mês] (não: em quanto tempo mato o projeto; porque "acabar em 60
   meses" depende de quantos lotes vendo por mês)
5. **Número total de lotes a vender:** total − permuta física
6. **Entrada de venda:** [%] (tipicamente 10–20% do preço; pode ser R$ fixo)
7. **Número de parcelas e prazo:** [nº × meses] — é comum 120 parcelas mensais = 10 anos
8. **Correção durante a obra (meses 1–18):** INCC + [qual juros?]
9. **Correção depois da obra (mês 19+):** IPCA + [qual juros?]

**Na planilha:**

```
Aba FLUXO, colunas:
  Mês | Lotes vendidos (acum.) | Lotes vendidos (novo mês)
      | Entrada capturada | Parcela #1 | Parcela #2 | ... | Parcela #120

Para Mês 1, Parcela #1 (mês seguinte ao mês de venda):
  = SE(lote vendido no mês -1, preço com correção, 0)

Para cada mês subsequente:
  = SE(lote vendido há N meses, parcela é devida, 0)
  com correção apropriada
```

É complexo porque uma venda no mês 5 gera 120 parcelas — temos de rastrear cada coorte de lotes
e calcular suas parcelas.

### Inadimplência, distrato, retomada

A Lei 13.786/2018 criou o regime de responsabilidade compartilhada do incorporador e do lojista
em relação à inadimplência do comprador de lote. Você precisa saber disso porque muda o fluxo:

- **Inadimplência por atraso:** o comprador não paga a parcela no prazo, mas promete pagar depois. Isso é um recebível atrasado, não é perda de receita.
- **Distrato:** o comprador não quer mais; você retoma o lote. Pela lei, o lojista (se o lote foi vendido por imobiliária) perde comissão, você retoma o lote.
- **Retomada e revenda:** o lote agora está sua de novo. Você vende outra vez (esperemos com menos desconto), e a parcela antiga é perdida.

**Na planilha — modelo conservador:**

```
Premissa: taxa de distrato + inadimplência crônica: [%]
  (historicamente, 3–10% conforme região e produto)

Para cada coorte de lotes (lotes vendidos no mês M):
  Lotes retomados (mês N): = coorte × taxa de distrato
  Esses lotes voltam para vender (fila de revenda)
  A receita de parcelas desses lotes é cancelada a partir do mês de distrato
```

É chato de modelar porque complica a fórmula de receita. Mas é realista e protege você de
surpresa.

### Tributação: lucro presumido vs. RET

**Atenção:** esta é uma seção onde a IA quebra. Nunca pergunte à IA qual regime tributário é
melhor para loteamento. A resposta vem errada e com confiança.

O que você precisa saber para a planilha:

- **Lucro Presumido:** base de cálculo = 32% da receita de venda (incorporação) ou 16% (loteamento). Alíquota: 15% + 9% (CSLL).
- **RET (Regime Especial Tributário de Incorporação Imobiliária):** 4% sobre a receita, ponto. Mas tem condições: terreno precisa ser de primeira aquisição, obra tem de estar dentro do prazo, não pode estar finalizada.
- **LSJP (Lei de Segurança Jurídica em Propriedade):** criou uma série de regimes. O tratamento para loteador não é claro nas cartilhas.

**Na sua planilha, o que você faz:**

```
Aba TRIBUTAÇÃO:
  Regime 1 (Lucro Presumido): [fórmulas de cálculo]
  Regime 2 (RET): [fórmulas de cálculo]
  Regime 3 (LSJP ou regime municipal): [PENDÊNCIA TRIBUTÁRIA]
  
  Comparação: qual regime devolve maior VPL/TIR?
  
  [Célula de aviso]: "Confirmar regime com contador antes de decisão"
```

Depois você leva essa planilha — com a tributação ainda vaga — para o seu contador, ele confirma
qual regime se aplica e preenche a célula de [PENDÊNCIA]. A partir daí, a tributação fica
correta.

### Painel de métricas — a hierarquia

No final de tudo, uma única aba que resume a decisão. Em ordem hierárquica:

```
Aba DECISÃO:

1º — Viabilidade de caixa:
    Exposição máxima: R$ [X]
    Caixa disponível: R$ [Y]
    Status: ✓ Viável  /  ✗ Quebrado

Se ✗ Quebrado, PARAR. O resto é irrelevante.

2º — Retorno de capital:
    Payback de caixa (quantos meses até lucro acumulado virar positivo): [M] meses
    Payback de TIR (equivalente): [N] anos

Se não tem linha de crédito e o payback é 8 anos, talvez desista já.

3º — Rentabilidade:
    TIR mensal: [X%]
    TIR anual: [Y%]
    VPL (descontado a [Z]% a.a.): R$ [V]

4º — Margem e proporções:
    Margem sobre VGV: [%]
    VGV: R$ [total]
    Custo total (infra + aquisição): R$ [total]
    Lucro bruto: R$ [diferença]

A lógica: 1º você não quebra em caixa. 2º você recebe seu capital de volta em prazo
razoável. 3º ele rende o bastante. 4º a margin é defensável (não tão gorda a ponto de
convocar competidor, não tão magra a ponto de quebrar em risco).
```

---

<a name="parte-5"></a>
## Parte 5 — Laboratório: seis exercícios progressivos

Aqui você aprende fazendo. O caso é fictício mas realista: um loteamento em Palmas, TO, que
encapsula as decisões comuns em loteamento de verdade.

### O caso: Santa Helena Residencial

**Localização:** Palmas, Tocantins (região de expansão urbana)

**Premissas básicas:**
- Gleba: 50 lotes
- Área média do lote: 360 m²
- Preço: R$ 450/m² de lote (preço de piso — lotes grandes)
- Preço médio por lote: R$ 162 mil
- VGV: 50 × R$ 162 mil = R$ 8.100.000
- Aquisição: R$ 4.050.000 (à vista, 50% do VGV — barganha forte)
- Infra: R$ 2.700.000 total (R$ 54 mil/lote) — valor parametrizado de Palmas
- Entrada de venda: 15%
- Parcelas: 120 mensais (10 anos), INCC durante obra, IPCA + 0,7% a.m. depois
- Velocidade de vendas: 1,5 lotes/mês (baseline; varia nos exercícios)
- Aprovação: 6 meses
- Obra: 12 meses (mês 7–18)
- Início de venda: mês 8 (obra ainda andando)
- Distrato/inadimplência: 5%
- Regime tributário: Lucro Presumido 16%

### Exercício 1: Construir o esqueleto

**Objetivo:** que você veja uma planilha do começo sendo estruturada, e saia com um modelo de
"alta nível" pronto antes de plugar número nenhum.

**Prompt para você usar:**

```
Vou construir uma planilha de viabilidade de um loteamento (50 lotes, Palmas).
Antes de colocar número, preciso que você me mostre a estrutura.

Que abas a planilha deve ter? Para cada aba, liste as linhas que ela precisa ter.
Não coloque valores; só estrutura.

Aqui está o que sei:
- Loteamento de 50 lotes, preço médio R$ 162 mil/lote
- Gleba comprada à vista por R$ 4 milhões
- Infra total de R$ 2,7 milhões (corre ao longo de 12 meses)
- Venda começa no mês 8, com 1,5 lotes/mês
- Entrada de venda: 15%, resto parcelado em 120 meses
- Regime: Lucro Presumido 16%
- Preciso rastrear: caixa, TIR, exposição máxima, payback

Estruture isso. Se faltar premissa, marque [FALTA: descrição].
```

**Critério de aprovação:** você tem uma aba de PREMISSAS, uma aba de ESTRUTURA DE FLUXO, uma aba
de CÁLCULOS, uma aba de RESULTADO, e uma aba de VALIDAÇÃO. Nenhuma com número.

### Exercício 2: Alimentar premissas e travar validação

**Objetivo:** colocar os números no lugar, construir os testes de amarração, e não deixar o
modelo quebrar silenciosamente quando você muda algo.

**Prompt para você usar:**

```
Agora vou colocar as premissas. Aqui estão:

[lista de 15 premissas: aquisição, infra, venda, juros, etc. — copie do case acima]

Reescreva a aba PREMISSAS com esses valores.

Depois, crie três testes na aba VALIDAÇÃO:
1. Caixa inicial + entradas = caixa final + saídas? (SE não, vermelha)
2. Nº de lotes no fluxo = 50 (ou 50 - permuta)? (SE não, vermelha)
3. A receita de mês N é derivada só das parcelas ativas daquele mês? Valide com somatório cruzado.

Para cada teste, mostre a fórmula. Se a fórmula quebrar por falta de referência, avise qual é a
referência que está faltando.
```

**Critério de aprovação:** três células verdes na aba VALIDAÇÃO (todos os testes passam).

### Exercício 3: Plantar um erro propositalmente e auditar

**Objetivo:** aprender a usar a IA para auditar de verdade.

**O que você faz:**

1. Pegue o arquivo do exercício 2 (tudo correto).
2. Abra a aba de FLUXO, procure uma célula de parcela no meio (digamos, mês 50).
3. Mude a fórmula dela para referência errada (por exemplo, mude `=XLOOKUP(mês, lista_de_meses, receita)` para `=XLOOKUP(mês+1, ...)`).
4. Salve.
5. Dê o arquivo para a IA com este prompt:

```
Aqui está uma planilha de viabilidade de loteamento. Ela tem um erro de fórmula. 

Seu trabalho: procure descasamentos entre as linhas, erros de período, referências que não
fecham, lógica quebrada. NÃO recalcule o TIR ou VPL. Procure só erro de estrutura.

Diga: (a) onde está o erro, (b) qual era a intenção da célula, (c) como corrigir.

Se não encontrar erro, diga "nenhum erro estrutural encontrado".
```

**Critério de aprovação:** a IA acha o erro. Se não acha, você vai para o prompt e muda — pede
para procurar em seções específicas (tipo "Aba FLUXO, linhas de receita parcelada"), ou pede
que ela mostre as fórmulas das células da linha 50 uma a uma.

### Exercício 4: Sensibilidade — tabela preço × velocidade

**Objetivo:** entender como TIR (e caixa) reagem a mudanças de premissa.

**Prompt:**

```
Tenho um loteamento com premissas base:
- Preço/lote: R$ 162 mil
- Velocidade: 1,5 lotes/mês

Quero uma tabela de sensibilidade 2D:
- Eixo X: velocidade de vendas (0,8 / 1,0 / 1,2 / 1,5 / 1,8 lotes/mês)
- Eixo Y: preço/lote (R$ 140k / R$ 150k / R$ 162k / R$ 180k / R$ 200k)
- Células: TIR mensal para cada combinação

Crie também uma tabela de exposição máxima de caixa (em vez de TIR) para as mesmas combinações.

Quando a exposição máxima fica negativa (maior que meu caixa de R$ 2 milhões), marque em
vermelho.
```

**Critério de aprovação:** duas tabelas 5×5 lado a lado; a tabela de caixa tem pelo menos 3–4
células vermelhas.

**Insights do resultado:** qual cenário mata o projeto — preço baixo ou velocidade baixa? Qual é
o mínimo de cada um para viabilidade?

### Exercício 5: Permuta vs. compra à vista

**Objetivo:** comparar dois cenários realistas de aquisição e ver qual TIR é ilusória.

**Prompt:**

```
Cenário A (base): Compra à vista de 50 lotes por R$ 4 milhões.

Cenário B: Permuta física — o dono do terreno aceita receber 8 lotes de volta em vez de caixa.
Você compra à vista só os outros 30% do valor (R$ 1.215.000) e oferece 8 lotes.

Agora rode os dois cenários:
- Cenário A: VGV = 50 × R$ 162k; caixa investido = R$ 4M + infra
- Cenário B: VGV = 42 × R$ 162k; caixa investido = R$ 1.2M + infra; receita é 42 lotes, não 50

Compare: TIR, VPL, exposição máxima, payback, margem sobre VGV.

Qual cenário tem maior TIR? Qual tem maior VPL? Por quê são diferentes?
```

**Critério de aprovação:** você entende por que TIR sobe em permuta (capital investido cai) e
VPL pode cair (receita cai mais). Ilusão estatística.

### Exercício 6: Red team — o sócio cético

**Objetivo:** que a IA ataque seu estudo como se fosse vendedor de gleba competidor.

**Prompt:**

```
Você é uma construtora concorrente. Alguém te mostrou esse estudo de viabilidade do loteamento
e disse "olha, é viável". Seu trabalho é QUEBRAR esse estudo.

Assuma que:
1. O preço de venda é otimista (5–10% acima da média regional)
2. A velocidade de 1,5 lotes/mês é média — mas em recessão cai para 0,8
3. O custo de infra foi orçado, não licitado — pode estar 15% abaixo do real

Com esses pressupostos:
- Qual é a TIR pessimista?
- Qual é a exposição máxima de caixa no cenário pessimista?
- Com essa exposição, eu consigo financiar?

Aponte os três maiores riscos da tese, em ordem de impacto na TIR.
```

**Critério de aprovação:** você tem uma resposta que mostra vulnerabilidades reais. Depois,
você decide se concorda com o red team ou se sua tese sobrevive ao ataque.

---

<a name="parte-6"></a>
## Parte 6 — Trilha de aprofundamento

Você acabou de fazer seis exercícios. Agora, em ordem, o que estudar para não virar dependente
de IA e desenvolver julgamento de verdade.

### Bloco 1: Fundamentos de fluxo de caixa (4–6 horas)

O que você precisa aprender que não sabia:
- **Diferença entre fluxo operacional, fluxo de investimento, fluxo de financiamento.** Em loteamento, você tem os três rodando em paralelo e descasados.
- **Valor do dinheiro no tempo, de verdade.** Não só a fórmula de VPL, mas **por que** o desconto funciona (oportunidade de reinvest, inflação, risco).
- **Fluxo de caixa vs. resultado contábil.** Uma parcela de venda gera resultado (receita) no mês de venda, mas caixa chega ao longo de 120 meses. Sua exposição é função do fluxo de caixa, não do resultado.

Começar com:
- Uma planilha manual de um loteamento pequeno (20 lotes, 60 meses) desenhada em papel. Rode VPL à mão. Veja como a mão treme.
- Depois bota no Excel e deixa a máquina fazer. A diferença entre "eu entendo a fórmula" e "eu calculi uma vez" é enorme.

### Bloco 2: Incerteza e sensibilidade (3–4 horas)

- **Que é análise de sensibilidade e por que não é previsão.** Sensibilidade te diz "se a velocidade cair 20%, TIR cai de 35% para 18%". Não diz "velocidade vai cair 20%".
- **Monte Carlo: simulação vs. ponto.** Quando a premissa é não-determinística (velocidade de vendas é aleatória, preço flutua), você não roda "caso base", "pessimista", "otimista". Você roda 10 mil simulações e vê a distribuição.
- **Correlação entre premissas.** Preço de venda e velocidade não são independentes: quando preço sobe, demanda cai. Modelar esses dois como if-then é melhor que modelar como independentes.

Ferramenta: Solver do Excel ou tabela de dados. Evite softwares que fazem tudo "por você" neste
stage — você precisa entender o que está sendo simulado.

### Bloco 3: Estrutura de capital e financiamento (2–3 horas)

- **Seu capital próprio vs. capital de terceiros.** Se o projeto precisa de R$ 4 milhões de caixa mínima, de onde sai? É seu dinheiro? É crédito? É ambos?
- **Custo do capital de terceiros.** Se você pega R$ 3 milhões de linha de crédito a 14% a.a., isso muda o VPL — você tem de descontar o custo financeiro do projeto.
- **Quando o financiamento muda a viabilidade.** Um projeto com TIR de 22% é viável se você tem capital próprio à vontade. Se você tem R$ 1 milhão e o projeto precisa de R$ 2 milhões de desembolso máximo, você precisa financiar R$ 1 milhão a uma taxa. Se a taxa for 12%, o VPL segue positivo. Se for 18%, o projeto morre.

Ferramenta: Planilha com duas abas, "cenário sem financiamento" e "cenário com financiamento".

### Bloco 4: Tributação (realmente existe) (2–3 horas, depois com contador)

Nesta seção, a IA é fraca. Um contador é forte. Sua tarefa é:

1. Ler a Lei 6.766 e Lei 13.786 (os textos, não resumos). Você não precisa ser expert, mas
   precisa conhecer a moldura legal de loteamento.
2. Perguntar para seu contador:
   - Qual é o regime tributário aplicável ao seu loteamento?
   - Qual é a alíquota efetiva?
   - A tributação reajeita o cronograma (pagamos no mês da venda ou no mês do recebimento)?
3. Montar a planilha com as respostas.

Ferramenta: Lei 6.766, Lei 13.786, uma consulta com contador (2–3 horas de papo, R$ 300–500
dependendo do profissional).

Não tente automatizar esta parte com IA. Simplesmente não dá. Delegue.

### Bloco 5: Leitura de casos reais (0 horas de "estudo", 8 horas de imersão)

Pegue dois loteamentos de verdade — um seu, um de alguém que você confia — e rode o exercício 3
em cada um: dê a planilha para a IA, peça que critique. Depois incorpore a crítica e recalcule.

Você vai aprender mais aqui do que em qualquer texto.

### Bloco 6: Limite de aprendizado — quando parar

Você aprendeu o bastante quando consegue:

1. **Detectar erro de fórmula vendo o resultado.** A IA diz TIR = 27%, você lê e pensa "se a
   velocidade é 1,5 lotes/mês e temos 50 lotes, o projeto acaba em 33 meses; TIR de 27% a.m.
   faria sentido se o faturamento fosse no mês 1, mas entra parcelado em 120 meses; deve estar
   menos". Depois confere a fórmula e acha o erro.

2. **Duvidar de premissa.** A IA ou um sócio diz "compramos a gleba por R$ 4 milhões". Você
   pergunta: "a gleba é de quantos lotes? Quantos lotes per capita? É preço in-the-money ou
   desconto de urgência?" — e só aí acredita.

3. **Reescrever o modelo sem ajuda.** Você senta, abre uma planilha em branco e monta a
   estrutura de novo, sem IA, sem olhar a anterior. Se conseguir em 2 horas e o modelo ficar
   correto, você aprendeu.

---

<a name="parte-7"></a>
## Parte 7 — Régua de confiança

Checklist para colar ao lado do monitor. Diz o que aceitar da IA de olhos fechados, o que
conferir por amostragem, e o que nunca, jamais aceitar.

### Verde — aceitar sem conferir

- ✓ Estrutura de abas proposta antes de plugar número
- ✓ Fórmula de `=XTIR(...)` ou `=XVPL(...)` (se o intervalo está correto)
- ✓ Crítica de lógica (tipo: "a fórmula referencia a aba errada")
- ✓ Ideias de sensibilidade ou cenário a explorar

### Amarelo — conferir por amostragem

- ⚠ Fórmula complexa com muitos SEs aninhados (recalcule à mão um caso)
- ⚠ TIR ou VPL devolvida sem fórmula visível (pegue a fórmula e reproduza no Excel)
- ⚠ Resposta de IA a pergunta de tributação (sempre repasse para contador)
- ⚠ Premissa que a IA "deduz" sem que você tenha dito (sempre confirme a origem)

### Vermelho — nunca aceitar

- ✗ TIR de cabeça, sem arquivo de planilha
- ✗ Resposta a "qual lote é melhor para permutar" (decisão comercial, não IA)
- ✗ Parecer jurídico ("a lei permite...?") — sempre advogado
- ✗ Número sem célula que você consegue clicar e ver de onde veio
- ✗ Tributação recomendada sem aviso de "confirme com contador"

### Aviso de escopo

Este documento não é recomendação de investimento. Não é parecer contábil. Não é parecer jurídico.
Os números usados aqui são fictícios e servem só para ensinar método.

Antes de qualquer decisão de compra de gleba ou lançamento de loteamento:

1. **Contador:** regime tributário, imposto, fluxo de caixa fiscal
2. **Advogado:** Lei 6.766, Lei 13.786, Lei de Segurança Jurídica, contrato de permuta, registro,
   responsabilidades
3. **Engenheiro:** orçamento de infra, cronograma realista, riscos de obra
4. **Mercado:** preço de lote, velocidade de vendas, comparável regional

Depois disso, use IA e planilha para sintetizar. Antes disso, IA e planilha são brincadeira.

---

**Fim do documento**

**Para próxima sessão:** o repositório agora tem duas frentes. A minuta jurídica fica em
`docs/`, com regra dos três formatos. Os estudos de viabilidade ficam em `estudos/`, formato
único (markdown), sem Artifact. Ver `CLAUDE.md` para detalhes.
