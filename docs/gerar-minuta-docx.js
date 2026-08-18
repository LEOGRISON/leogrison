const d = require('docx');
const fs = require('fs');
const {Document,Packer,Paragraph,TextRun,AlignmentType,Table,TableRow,TableCell,WidthType,ShadingType,BorderStyle,Footer} = d;

const BODY="Cambria", SZ=19, ACC="1F5140", INK="141C17", MUT="4A5850";
const sp=(a,b)=>({before:a,after:b});
const R=(runs,o={})=>new Paragraph({
  spacing:o.spacing||sp(0,60), alignment:o.align||AlignmentType.JUSTIFIED, indent:o.indent, keepNext:o.keepNext,
  children:runs.map(([t,f={}])=>new TextRun({text:t,font:BODY,size:f.size||o.size||SZ,color:f.color||o.color||INK,bold:f.bold,italics:f.italics}))
});
const H=(text)=>new Paragraph({
  spacing:sp(150,60), keepNext:true,
  children:[new TextRun({text:text.toUpperCase(),font:BODY,size:18,bold:true,color:ACC,characterSpacing:22})]
});
const RULE=(before=0,after=120)=>new Paragraph({spacing:sp(before,after),
  border:{bottom:{style:BorderStyle.SINGLE,size:6,color:"C3CDC5",space:1}},children:[]});

const W=[1080,2620,6540];
const cell=(children,w)=>new TableCell({width:{size:w,type:WidthType.DXA},
  margins:{top:70,bottom:70,left:90,right:90},children});
const metaRow=(n,titulo,valor,desc)=>new TableRow({children:[
  cell([R([[n,{bold:true,color:ACC,size:18}]],{align:AlignmentType.LEFT,spacing:sp(0,0)})],W[0]),
  cell([R([[titulo,{bold:true}]],{align:AlignmentType.LEFT,spacing:sp(0,20)}),
        R([[valor,{bold:true,color:ACC,size:19}]],{align:AlignmentType.LEFT,spacing:sp(0,0)})],W[1]),
  cell([R([[desc,{size:18,color:MUT}]],{spacing:sp(0,0)})],W[2])
]});

const doc = new Document({
  creator:"Grupo Grison", title:"Distribuicao de lucros por metas",
  sections:[{
    properties:{page:{size:{width:12240,height:15840},margin:{top:960,right:1000,bottom:860,left:1000,footer:420}}},
    footers:{default:new Footer({children:[
      new Paragraph({spacing:sp(0,0),border:{top:{style:BorderStyle.SINGLE,size:6,color:"C3CDC5",space:6}},children:[
        new TextRun({text:"Documento de trabalho para discussão familiar. Não constitui parecer jurídico e não substitui a análise do advogado da família e do tributarista, que devem validar a redação antes de qualquer assinatura.",font:BODY,size:15,color:MUT,italics:true})]})]})},
    children:[

    new Paragraph({spacing:sp(0,60),border:{top:{style:BorderStyle.SINGLE,size:18,color:ACC,space:6}},children:[
      new TextRun({text:"MINUTA PARA DISCUSSÃO INTERNA  ·  VERSÃO 2.0  ·  NÃO ASSINADA",font:BODY,size:15,color:MUT,characterSpacing:18})]}),
    new Paragraph({spacing:sp(60,80),children:[
      new TextRun({text:"Distribuição de lucros por metas, não por prazo",font:BODY,size:31,bold:true,color:INK})]}),
    R([["Proposta de redação para uma cláusula do acordo de sócios da holding familiar: nenhum sócio retira lucro enquanto o grupo não estiver com dívida zero, caixa formado e a moradia da mãe garantida — sem prazo mínimo e sem prazo máximo. Os pais, usufrutuários, estão fora desta regra.",{size:18,color:MUT}]],{spacing:sp(0,30)}),
    RULE(60,140),

    H("Por que meta e não prazo"),
    R([["A trava por prazo desestimula o esforço. ",{bold:true}],["Se todos sabem que nada será distribuído antes de cinco anos, o trabalho de hoje não muda nada, e a energia da família vai para projetos onde o retorno aparece antes — justamente o que a empresa não pode perder agora."]]),
    R([["A trava por meta é mais rigorosa, não mais frouxa. ",{bold:true}],["Pela regra de prazo, no dia seguinte ao quinto ano abre-se a distribuição ainda que as dívidas estejam de pé e o caixa vazio. Pela regra de metas isso é impossível: enquanto houver dívida em aberto, ninguém retira — em cinco, sete ou dez anos."]]),
    R([["O prazo deixa de ser premissa e vira consequência. ",{bold:true}],["Se levar cinco anos, o efeito prático terá sido idêntico ao da regra original. Se conseguir em três, a antecipação é o prêmio de quem trabalhou."]]),
    R([["A meta se audita; o prazo não mede nada. ",{bold:true}],["Dívida zero, saldo em caixa e reserva são fatos verificáveis em balanço, certidão e ata — sem discussão futura sobre quem pode retirar o quê."]]),

    H("As três metas — cumulativas"),
    new Table({columnWidths:W,rows:[
      metaRow("META I","Dívida zero","100% do perímetro","Saldo zero de tudo — financiamentos, empréstimos, fornecedores, tributos, multas, custas e emolumentos. Só não conta o adiantamento do loteador."),
      metaRow("META II","Caixa-pulmão","R$ 2.000.000,00","Caixa livre consolidado, sustentado por dois balancetes consecutivos. Impede que a próxima parcela vire nova dívida."),
      metaRow("META III","Moradia da usufrutuária","R$ 1.600.000,00","Reserva vinculada para o imóvel residencial da Sra. Elman, ou imóvel já adquirido. Vem antes de qualquer retirada dos filhos."),
    ]}),
    R([["Quem está sujeito à regra: ",{bold:true,size:18}],["filhos, cotistas e sócios admitidos no futuro, inclusive colaboradores em partnership. ",{size:18,color:MUT}],["Os usufrutuários Wilson e Elman estão fora: os frutos lhes pertencem por direito e seguem pagos integralmente, sem limite e sem condição.",{size:18,color:MUT,bold:true}]],{spacing:sp(110,0)}),

    H("Minuta — Cláusula [•]: destinação de resultados e metas de liberação"),

    R([["1. Princípio e alcance. ",{bold:true}],["A distribuição de lucros, dividendos, juros sobre capital próprio ou qualquer outra forma de retorno do capital aos "],["Sócios Sujeitos à Retenção",{bold:true}],[" não é regida por prazo, mas pelo cumprimento cumulativo das três Metas do item 3. São Sócios Sujeitos à Retenção os filhos nu-proprietários, os demais cotistas e quaisquer sócios admitidos no futuro, inclusive colaboradores em regime de partnership. "],["Esta Cláusula não se aplica aos usufrutuários",{bold:true}],[" Wilson Grison e Elman M. Coelho Grison: os frutos das participações gravadas com usufruto lhes pertencem por direito próprio e são pagos integralmente, sem limite, prazo ou condição."]]),

    R([["2. Definições.",{bold:true}]],{spacing:sp(50,50)}),
    R([["2.1  Perímetro de Endividamento do Grupo. ",{bold:true}],["Todas as obrigações pecuniárias originadas nos negócios conduzidos, liderados ou orientados pelo Sr. Wilson Grison, "],["independentemente da pessoa em cujo nome tenham sido contraídas",{bold:true}],[", abrangendo financiamentos e empréstimos bancários, dívidas com fornecedores e prestadores de serviço, tributos de qualquer esfera, multas e autos de infração, custas e emolumentos cartorários, parcelamentos e transações, e obrigações decorrentes de garantias prestadas. Alcança as dívidas da Grison e Cia Ltda., as pessoais de Wilson e de Elman M. Coelho Grison como devedores, avalistas, fiadores ou coobrigados, e as das sociedades por eles constituídas, controladas ou dirigidas de fato, ainda que em nome de filhos, cônjuge ou terceiros — EMC Grison, Bioaçaí, ML Serviços Agrícolas, Natyrê, Tiúba 2 SPE e Patrimônio Digital. Obrigação não listada no Anexo I que se enquadre nesta definição integra o Perímetro."]],{indent:{left:200}}),
    R([["2.2  Não integram o Perímetro ",{bold:true}],["as dívidas estritamente particulares de qualquer sócio, sem relação com os negócios do grupo, pelas quais cada um responde por conta própria."]],{indent:{left:200}}),
    R([["2.3  Adiantamento Santa Helena. ",{bold:true}],["Valores antecipados pelo loteador parceiro por conta de resultados futuros da parceria, recompostos por retenção, pelo próprio loteador, dos repasses ao grupo. "],["Está expressamente excluído do Perímetro",{bold:true}],[" e não conta para a Meta I, por constituir obrigação de liquidação automática por compensação contratual, independente de desembolso ou de ato de gestão dos sócios. Consta do Anexo I apenas para informação."]],{indent:{left:200}}),
    R([["2.4  Dívida Quitada. ",{bold:true}],["Aquela cujo saldo é zero e cuja exigibilidade se extinguiu em face de todos os coobrigados — por pagamento, deságio, transação, novação, decisão transitada em julgado, prescrição ou assunção liberatória com exoneração dos garantidores. É indiferente o valor desembolsado, mas "],["parcelamento em curso, ainda que rigorosamente adimplente, não é dívida quitada",{bold:true}],["."]],{indent:{left:200}}),
    R([["2.5  Caixa Mínimo. ",{bold:true}],["Caixa e equivalentes de livre movimentação da Companhia e controladas, excluídos os recursos vinculados à Meta III, os de terceiros e os valores bloqueados ou de destinação vinculada."]],{indent:{left:200}}),

    R([["3. As Metas — cumulativas e não fracionáveis. ",{bold:true}],["(I) "],["Dívida zero",{bold:true}],[": saldo zero de 100% das obrigações do Perímetro, ressalvado o Adiantamento Santa Helena; (II) Caixa Mínimo igual ou superior a R$ 2.000.000,00, mantido por dois balancetes mensais consecutivos; (III) Reserva Moradia igual ou superior a R$ 1.600.000,00 destinada ao imóvel residencial da Sra. Elman, ou imóvel já adquirido e quitado. Os valores das Metas II e III são corrigidos anualmente pelo IPCA. O cumprimento parcial, ainda que de duas Metas, não autoriza distribuição parcial, proporcional ou antecipada."]],{spacing:sp(50,60)}),
    R([["3.1  Comprovação objetiva da Meta I. ",{bold:true}],["A Meta I só se tem por cumprida mediante, cumulativamente: certidões negativas — ou positivas com efeito de negativas — federais, estaduais e municipais da Grison e Cia Ltda. e das demais pessoas do Perímetro; ausência de protesto; ausência de anotação restritiva em Serasa, SPC ou equivalente relativa a obrigação do Perímetro; e baixa dos gravames que as garantiam. Anotação decorrente de dívida particular de sócio não impede o cumprimento da Meta."]],{indent:{left:200}}),

    R([["4. Retenção. ",{bold:true}],["Até a Data de Liberação é vedada aos Sócios Sujeitos à Retenção qualquer distribuição de lucros, dividendos, juros sobre capital próprio, redução de capital com restituição, resgate ou reembolso, ressalvado o item 5. Os lucros que lhes caberiam são destinados a reserva fundamentada em orçamento de capital aprovado anualmente em Assembleia (art. 196 da Lei 6.404/76), obrigando-se eles a votar favoravelmente a essa destinação; o Estatuto fixará o dividendo obrigatório em percentual compatível com esta Cláusula (art. 202). A retenção não alcança os frutos devidos aos usufrutuários, pagos normalmente."]]),

    R([["5. Exceções. ",{bold:true}],["Não configuram distribuição vedada: (a) as parcelas do ITCMD da doação e os custos dos atos societários e registrais correlatos; (b) o pró-labore, em valor de mercado e aprovado em Assembleia, dos sócios que exerçam função na Companhia ou em suas controladas; (c) o reembolso de despesas comprovadas no interesse da Companhia; (d) a devolução de mútuo formalizado com juros de mercado e encargos recolhidos. Os valores de (b) constam do Anexo II e sua majoração não pode ter por efeito substituir a distribuição vedada."]]),

    R([["6. Vedação à distribuição disfarçada. ",{bold:true}],["Durante a retenção é vedado o pagamento de despesas pessoais de Sócio Sujeito à Retenção pela Companhia ou controladas, o mútuo a sócio ou parte relacionada sem contrato e juros de mercado, a contratação fora de mercado, o uso de bem social sem contrapartida e o desvio de oportunidade de negócio. As sociedades manterão contabilidades segregadas e contratos escritos entre si. A violação obriga à restituição integral e corrigida."]]),

    R([["7. Verificação. ",{bold:true}],["A administração apresenta Relatório de Progresso das Metas semestralmente. Entendendo-as cumpridas, convoca Assembleia instruída com balancete de até 30 dias, declaração do contador, comprovação da extinção de cada obrigação e as certidões e consultas do item 3.1, que declara em ata a Data de Liberação. Qualquer sócio pode requerer, às suas expensas, verificação independente."]]),

    R([["8. Efeitos e restabelecimento da trava. ",{bold:true}],["Aberta a distribuição, ela não pode reduzir o Caixa Mínimo abaixo do valor da Meta II — que deixa de ser meta e passa a ser piso permanente — e a Reserva Moradia segue indisponível até a aquisição. A distribuição fica automaticamente suspensa se for revelada obrigação do Perímetro anterior à declaração e dela não constante, ou se o caixa cair abaixo do piso."]]),

    R([["9. Ausência de prazo. ",{bold:true}],["Não há prazo mínimo nem máximo, e nenhum decurso de tempo produz liberação automática. Decorridos 60 meses sem cumprimento integral, a Assembleia revisa o plano de execução e o cronograma; a revisão não flexibiliza nem dispensa as Metas, o que dependerá de deliberação específica pelo quórum de alteração do Acordo."]]),

    R([["Anexos. ",{bold:true}],["I — Mapa de Dívidas do Perímetro (credor, natureza, situação processual, garantias, coobrigados e status), com o Adiantamento Santa Helena destacado como item excluído. II — Pró-labore aprovado."]]),

    H("Dois pontos para o advogado antes de fechar"),
    R([["1.  Parcelamento fiscal e a Meta I. ",{bold:true}],["Como parcelamento em curso não conta como quitado (item 2.4), uma eventual adesão à transação da PGFN em 120 meses adia a Meta I até a quitação do saldo, mesmo com certidão limpa e sem restrição cadastral. É consequência assumida da regra de dívida zero — vale confirmar se é o efeito desejado."]],{indent:{left:200}}),
    R([["2.  Mecanismo de retenção seletiva. ",{bold:true}],["Reter o lucro de alguns sócios e pagar os frutos aos usufrutuários exige base estatutária: dividendo obrigatório compatível, reserva com orçamento de capital e obrigação de voto no acordo — ou classes distintas de ações para os sócios em partnership. Definir com o advogado."]],{indent:{left:200}}),
  ]}]
});

Packer.toBuffer(doc).then(b=>{fs.writeFileSync("/home/user/leogrison/docs/Minuta-Distribuicao-por-Metas.docx",b);console.log("ok",b.length)});
