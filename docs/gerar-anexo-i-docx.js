const d = require('docx');
const fs = require('fs');
const {Document,Packer,Paragraph,TextRun,AlignmentType,Table,TableRow,TableCell,WidthType,BorderStyle,Footer,LineRuleType} = d;

const F="Times New Roman", SZ=22, INK="000000", MUT="444444";
const LS={line:264,lineRule:LineRuleType.AUTO};

// numbered clause item: "1.1." + text runs
const item=(n,runs,o={})=>new Paragraph({
  spacing:{...LS,before:o.before||0,after:o.after===undefined?130:o.after},
  alignment:AlignmentType.JUSTIFIED,
  indent:o.indent,
  children:[
    ...(n?[new TextRun({text:n+"  ",font:F,size:SZ,bold:true,color:INK})]:[]),
    ...runs.map(([t,f={}])=>new TextRun({text:t,font:F,size:f.size||SZ,color:f.color||INK,bold:f.bold,italics:f.italics}))
  ]
});
const sec=(n,t)=>new Paragraph({
  spacing:{...LS,before:225,after:130}, keepNext:true,
  children:[new TextRun({text:n+"  "+t.toUpperCase(),font:F,size:SZ,bold:true,color:INK})]
});
const plain=(runs,o={})=>item(null,runs,o);

// example table
const TW=[3400,3100,2600];
const c=(txt,w,o={})=>new TableCell({width:{size:w,type:WidthType.DXA},
  margins:{top:60,bottom:60,left:100,right:100},
  children:[new Paragraph({spacing:{line:240,lineRule:LineRuleType.AUTO,after:0},
    alignment:o.right?AlignmentType.RIGHT:AlignmentType.LEFT,
    children:[new TextRun({text:txt,font:F,size:21,bold:o.bold,color:INK})]})]});
const row=(a,b,v,o={})=>new TableRow({tableHeader:!!o.head,cantSplit:true,children:[c(a,TW[0],o),c(b,TW[1],o),c(v,TW[2],{...o,right:true})]});

const doc = new Document({
  creator:"Grupo Grison", title:"Anexo I - Criterios de Destinacao de Resultados",
  sections:[{
    properties:{page:{size:{width:12240,height:15840},margin:{top:1040,right:1220,bottom:860,left:1220,footer:400}}},
    footers:{default:new Footer({children:[
      new Paragraph({spacing:{after:0},border:{top:{style:BorderStyle.SINGLE,size:6,color:"AAAAAA",space:6}},children:[
        new TextRun({text:"Minuta de trabalho, versão 3.0. Não constitui parecer jurídico; sujeita à revisão do advogado da família e de tributarista antes da assinatura.",font:F,size:16,color:MUT,italics:true})]})]})},
    children:[

    new Paragraph({spacing:{after:80},alignment:AlignmentType.CENTER,children:[
      new TextRun({text:"ANEXO I AO ACORDO DE SÓCIOS",font:F,size:24,bold:true,color:INK})]}),
    new Paragraph({spacing:{after:340},alignment:AlignmentType.CENTER,children:[
      new TextRun({text:"CRITÉRIOS DE DESTINAÇÃO DE RESULTADOS",font:F,size:24,bold:true,color:INK})]}),

    // ---- nota curta ----
    new Paragraph({spacing:{...LS,after:150},alignment:AlignmentType.JUSTIFIED,children:[
      new TextRun({text:"Nota preliminar — a lógica do anexo (não integra o texto contratual)",font:F,size:21,bold:true,italics:true,color:INK})]}),
    plain([["A distribuição não é travada por prazo, e sim por metas. Prazo fixo desestimula: se ninguém recebe nada antes de cinco anos, o esforço de hoje não altera nada. A meta é mais rigorosa, porque nenhum decurso de tempo libera a distribuição enquanto houver dívida em aberto, e premia quem resolver antes.",{size:21}]],{after:130}),
    plain([["Quanto ao imóvel da usufrutuária: o custo deve recair apenas sobre Leonardo e Bárbara, filhos biológicos e herdeiros dela. Como o dinheiro sai da Companhia, que pertence aos quatro, distribui-se um montante bruto calibrado para que a fatia dos dois cubra exatamente o imóvel e os demais recebam em dinheiro o equivalente à sua participação. Daí a divisão por 74,62%, soma das participações de Leonardo e Bárbara — adotada com duas casas, e não pela fração exata de 74,625%, para que o número do contrato seja o da planilha.",{size:21}]],{after:130}),
    plain([["Renumeração: o Mapa de Dívidas passa a Anexo II e o pró-labore aprovado, a Anexo III.",{size:21}]],{after:300}),

    // ---- 1 ----
    sec("1.","Alcance"),
    item("1.1.",[["São Sócios Sujeitos à Retenção os filhos nu-proprietários, os demais cotistas e quaisquer sócios admitidos no futuro, inclusive colaboradores em regime de partnership."]]),
    item("1.2.",[["Este Anexo não se aplica aos usufrutuários Wilson Grison e Elman M. Coelho Grison. Os frutos das participações gravadas com usufruto lhes pertencem por direito próprio e são pagos integralmente, sem limite, prazo ou condição."]]),

    // ---- 2 ----
    sec("2.","Perímetro de Endividamento do Grupo"),
    item("2.1.",[["Integram o Perímetro as obrigações pecuniárias, vencidas ou vincendas, originadas nos negócios conduzidos, liderados ou orientados por Wilson Grison, "],["independentemente da pessoa em cujo nome tenham sido contraídas",{bold:true}],[": financiamentos e empréstimos bancários, dívidas com fornecedores e prestadores de serviço, tributos de qualquer esfera, multas e autos de infração, custas e emolumentos cartorários, parcelamentos e transações, e obrigações decorrentes de garantias prestadas."]]),
    item("2.2.",[["Alcança as dívidas da Grison e Cia Ltda., as pessoais de Wilson Grison e de Elman M. Coelho Grison como devedores, avalistas, fiadores ou coobrigados, e as das sociedades por eles constituídas, controladas ou dirigidas de fato, ainda que em nome de filhos, cônjuge ou terceiros, incluindo EMC Grison, Bioaçaí, ML Serviços Agrícolas, Natyrê, Tiúba 2 SPE e Patrimônio Digital. Obrigação não relacionada no Anexo II que se enquadre nesta definição integra o Perímetro."]]),
    item("2.3.",[["Não integram o Perímetro as dívidas estritamente particulares de qualquer sócio, sem relação com os negócios do grupo."]]),
    item("2.4.",[["Não integra o Perímetro o adiantamento recebido do loteador parceiro (Santa Helena), recomposto por retenção, pelo próprio loteador, dos repasses devidos ao grupo. Sua liquidação é automática e independe de desembolso ou de ato de gestão dos sócios."]]),
    item("2.5.",[["Dívida Quitada é aquela cujo saldo é zero e cuja exigibilidade se extinguiu perante todos os coobrigados, por pagamento, deságio, transação, novação, decisão transitada em julgado, prescrição ou assunção liberatória com exoneração dos garantidores. É indiferente o valor desembolsado. "],["Parcelamento em curso, ainda que adimplente, não é Dívida Quitada.",{bold:true}]]),

    // ---- 3 ----
    sec("3.","Metas de liberação"),
    item("3.1.",[["A distribuição aos Sócios Sujeitos à Retenção fica condicionada ao cumprimento cumulativo das seguintes Metas:"]],{after:110}),
    item("(a)",[["Meta I — Dívida zero. ",{bold:true}],["Quitação de 100% das obrigações do Perímetro."]],{indent:{left:420},after:90}),
    item("(b)",[["Meta II — Caixa mínimo. ",{bold:true}],["Caixa e equivalentes de livre movimentação iguais ou superiores a R$ 2.000.000,00 (dois milhões de reais), mantidos por dois balancetes mensais consecutivos."]],{indent:{left:420},after:90}),
    item("(c)",[["Meta III — Imóvel da usufrutuária. ",{bold:true}],["Disponibilidade, em reserva vinculada, do montante bruto apurado na forma do item 4.3, correspondente a R$ 2.144.197,27 (dois milhões, cento e quarenta e quatro mil, cento e noventa e sete reais e vinte e sete centavos) para imóvel mobiliado de R$ 1.600.000,00 — ou operação já concluída."]],{indent:{left:420}}),
    item("3.2.",[["A Meta I somente se tem por cumprida mediante, cumulativamente: certidões negativas, ou positivas com efeito de negativas, federais, estaduais e municipais da Grison e Cia Ltda. e das demais pessoas do Perímetro; ausência de protesto; ausência de anotação restritiva em Serasa, SPC ou equivalente; e baixa dos gravames. Anotação decorrente de dívida particular de sócio não impede o cumprimento da Meta."]]),
    item("3.3.",[["Os valores das Metas II e III são corrigidos anualmente pelo IPCA a contar da assinatura do Acordo. O cumprimento parcial não autoriza distribuição parcial, proporcional ou antecipada."]]),

    // ---- 4 ----
    sec("4.","Imóvel da usufrutuária e equalização entre os sócios"),
    item("4.1.",[["O imóvel residencial destinado à usufrutuária Elman M. Coelho Grison, de sua livre escolha, será adquirido mobiliado pela Companhia e registrado em nome dela, livre e desembaraçado."]]),
    item("4.2.",[["O custo econômico da aquisição recai exclusivamente sobre Leonardo Grison e Bárbara Grison, filhos biológicos e herdeiros da usufrutuária. Wellington Grison e Marcelo Grison não concorrem para esse custeio e recebem em dinheiro o correspondente à sua participação no montante bruto."]]),
    item("4.3.",[["Fórmula. ",{bold:true}],["O montante bruto da operação corresponde ao valor do imóvel mobiliado dividido por 74,62%, percentual que corresponde à soma das participações de Leonardo (37,31%) e Bárbara (37,31%)."]]),
    new Paragraph({spacing:{...LS,after:170},alignment:AlignmentType.CENTER,children:[
      new TextRun({text:"Montante Bruto  =  Valor do Imóvel Mobiliado  ÷  74,62%",font:F,size:22,bold:true,color:INK})]}),
    item("4.4.",[["Apurado o montante bruto, a operação se decompõe em:"]],{after:110}),
    item("(a)",[["37,31% a Bárbara Grison e 37,31% a Leonardo Grison, satisfeitos em bem, pela aquisição do imóvel;"]],{indent:{left:420},after:90}),
    item("(b)",[["12,44% a Wellington Grison, em dinheiro;"]],{indent:{left:420},after:90}),
    item("(c)",[["12,44% a Marcelo Grison, em dinheiro;"]],{indent:{left:420},after:90}),
    item("(d)",[["0,5% a Leonardo Grison, em dinheiro, correspondente à sua participação originária na Grison e Cia Ltda."]],{indent:{left:420}}),
    item("4.5.",[["A aquisição do imóvel e os pagamentos em dinheiro são simultâneos, constituem execução da Meta III e não configuram distribuição vedada pelo item 5."]]),
    item("4.6.",[["Os percentuais deste item aplicam-se exclusivamente a esta operação e não estabelecem critério de rateio para as demais distribuições."]]),
    item("4.7.",[["O enquadramento tributário da aquisição pela Companhia e da transferência ao patrimônio da usufrutuária será definido previamente com o tributarista, e seu custo integra o montante bruto."]]),
    item("4.8.",[["Aplicação da fórmula a imóvel mobiliado de R$ 1.600.000,00, a título ilustrativo:"]],{after:130}),

    new Table({columnWidths:TW,rows:[
      row("Sócio","Participação","Valor (R$)",{bold:true,head:true}),
      row("Bárbara Grison","37,31% — em bem","800.000,00"),
      row("Leonardo Grison","37,31% — em bem","800.000,00"),
      row("Wellington Grison","12,44% — em dinheiro","266.738,14"),
      row("Marcelo Grison","12,44% — em dinheiro","266.738,14"),
      row("Leonardo Grison","0,5% — em dinheiro","10.720,99"),
      row("Montante bruto","100%","2.144.197,27",{bold:true}),
    ]}),

    // ---- 5 ----
    sec("5.","Retenção"),
    item("5.1.",[["Até a Data de Liberação é vedada aos Sócios Sujeitos à Retenção qualquer distribuição de lucros, dividendos, juros sobre capital próprio, redução de capital com restituição, resgate ou reembolso, ressalvados os itens 4 e 6."]]),
    item("5.2.",[["A retenção não alcança os frutos devidos aos usufrutuários, pagos na forma do item 1.2."]]),
    item("5.3.",[["Os lucros que caberiam aos Sócios Sujeitos à Retenção são destinados a reserva fundamentada em orçamento de capital aprovado anualmente em Assembleia (art. 196 da Lei 6.404/76), obrigando-se eles a votar favoravelmente a essa destinação. O Estatuto fixará o dividendo obrigatório em percentual compatível com este Anexo (art. 202)."]]),

    // ---- 6 ----
    sec("6.","Exceções"),
    item("6.1.",[["Não configuram distribuição vedada: (a) as parcelas do ITCMD da doação das participações e os custos dos atos societários e registrais correlatos; (b) o pró-labore, em valor de mercado aprovado em Assembleia, dos sócios que exerçam função na Companhia ou em suas controladas; (c) o reembolso de despesas comprovadas no interesse da Companhia; (d) a devolução de mútuo formalizado com juros de mercado e encargos recolhidos."]]),
    item("6.2.",[["Durante a retenção é vedado o pagamento de despesas pessoais de Sócio Sujeito à Retenção pela Companhia ou controladas, o mútuo a sócio ou parte relacionada sem contrato e juros de mercado, a contratação fora de mercado e o desvio de oportunidade de negócio. A violação obriga à restituição integral e corrigida."]]),
    item("6.3.",[["As sociedades do grupo manterão contabilidades segregadas e formalizarão por contrato escrito as relações entre si."]]),

    // ---- 7 ----
    sec("7.","Verificação e Data de Liberação"),
    item("7.1.",[["A administração apresentará relatório semestral de progresso. Entendendo as Metas cumpridas, convocará Assembleia instruída com balancete de até 30 dias, declaração do contador, comprovação da extinção de cada obrigação e as certidões e consultas do item 3.2."]]),
    item("7.2.",[["A Assembleia declarará em ata a Data de Liberação. Qualquer sócio pode requerer, às suas expensas, verificação independente antes da deliberação."]]),
    item("7.3.",[["A partir da Data de Liberação, a distribuição observa a deliberação da Assembleia e o Estatuto, vedada a redução do caixa abaixo do valor da Meta II, que passa a ser piso permanente."]]),
    item("7.4.",[["A distribuição fica automaticamente suspensa se for revelada obrigação do Perímetro anterior à declaração e dela não constante, ou se o caixa cair abaixo do piso, até o restabelecimento da condição."]]),

    // ---- 8 ----
    sec("8.","Ausência de prazo"),
    item("8.1.",[["Não há prazo mínimo nem máximo para o cumprimento das Metas, e nenhum decurso de tempo produz liberação automática."]]),
    item("8.2.",[["Decorridos 60 meses sem cumprimento integral, a Assembleia revisará o plano de execução. A revisão não flexibiliza nem dispensa as Metas, o que dependerá de deliberação específica pelo quórum de alteração do Acordo."]],{after:0}),
  ]}]
});

Packer.toBuffer(doc).then(b=>{fs.writeFileSync("/home/user/leogrison/docs/Anexo-I-Destinacao-de-Resultados.docx",b);console.log("ok",b.length)});
