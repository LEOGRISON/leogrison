const d = require('docx');
const fs = require('fs');
const {Document,Packer,Paragraph,TextRun,HeadingLevel,AlignmentType,Table,TableRow,TableCell,WidthType,ShadingType,BorderStyle,Footer} = d;

const BODY="Cambria", SZ=19, ACC="1F5140", INK="141C17", MUT="4A5850";

const sp=(a,b)=>({before:a,after:b});
const P=(text,o={})=>new Paragraph({
  spacing:o.spacing||sp(0,60), alignment:o.align||AlignmentType.JUSTIFIED,
  indent:o.indent, keepNext:o.keepNext,
  children:[new TextRun({text,font:BODY,size:o.size||SZ,color:o.color||INK,bold:o.bold,italics:o.italics})]
});
// rich paragraph: array of [text, {bold,italics,color,size}]
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

const cell=(children,w,shade)=>new TableCell({
  width:{size:w,type:WidthType.DXA},
  shading:shade?{type:ShadingType.CLEAR,fill:shade,color:"auto"}:undefined,
  margins:{top:70,bottom:70,left:90,right:90},
  children
});
const W=[1080,2620,6540];
const metaRow=(n,titulo,valor,desc)=>new TableRow({children:[
  cell([R([[n,{bold:true,color:ACC,size:18}]],{align:AlignmentType.LEFT,spacing:sp(0,0)})],W[0]),
  cell([R([[titulo,{bold:true}]],{align:AlignmentType.LEFT,spacing:sp(0,20)}),
        R([[valor,{bold:true,color:ACC,size:19}]],{align:AlignmentType.LEFT,spacing:sp(0,0)})],W[1]),
  cell([R([[desc,{size:19,color:MUT}]],{spacing:sp(0,0)})],W[2])
]});

const doc = new Document({
  creator:"Grupo Grison", title:"Distribuicao de lucros por metas",
  sections:[{
    properties:{page:{size:{width:12240,height:15840},margin:{top:960,right:1000,bottom:860,left:1000,footer:420}}},
    footers:{default:new Footer({children:[
      new Paragraph({spacing:sp(0,0),border:{top:{style:BorderStyle.SINGLE,size:6,color:"C3CDC5",space:6}},children:[
        new TextRun({text:"Documento de trabalho para discussão familiar. Não constitui parecer jurídico e não substitui a análise do advogado da família e do tributarista, que devem validar a redação antes de qualquer assinatura.",font:BODY,size:15,color:MUT,italics:true})]})]})},
    children:[

    // ---------- cabeçalho ----------
    new Paragraph({spacing:sp(0,60),border:{top:{style:BorderStyle.SINGLE,size:18,color:ACC,space:6}},children:[
      new TextRun({text:"MINUTA PARA DISCUSSÃO INTERNA  ·  VERSÃO 1.0  ·  NÃO ASSINADA",font:BODY,size:15,color:MUT,characterSpacing:18})]}),
    new Paragraph({spacing:sp(60,80),children:[
      new TextRun({text:"Distribuição de lucros por metas, não por prazo",font:BODY,size:31,bold:true,color:INK})]}),
    R([["Proposta de redação para uma cláusula do acordo de sócios da holding familiar: nenhum filho retira lucro enquanto as dívidas do grupo não estiverem quitadas, o caixa formado e a moradia da mãe garantida — sem prazo mínimo e sem prazo máximo. Escopo limitado à destinação de resultados.",{size:18,color:MUT}]],{spacing:sp(0,30)}),
    RULE(60,140),

    // ---------- por que ----------
    H("Por que meta e não prazo"),
    R([["A trava por prazo desestimula o esforço. ",{bold:true}],["Se todos sabem que nada será distribuído antes de cinco anos, o trabalho de hoje não muda nada, e a energia da família vai para projetos onde o retorno aparece antes — justamente o que a empresa não pode perder agora."]]),
    R([["A trava por meta é mais rigorosa, não mais frouxa. ",{bold:true}],["Pela regra de prazo, no dia seguinte ao quinto ano abre-se a distribuição ainda que as dívidas estejam de pé e o caixa vazio. Pela regra de metas isso é impossível: enquanto houver dívida em aberto, ninguém retira — em cinco, sete ou dez anos."]]),
    R([["O prazo deixa de ser premissa e vira consequência. ",{bold:true}],["Se levar cinco anos, o efeito prático terá sido idêntico ao da regra original. Se conseguir em três, a antecipação é o prêmio de quem trabalhou."]]),
    R([["A meta se audita; o prazo não mede nada. ",{bold:true}],["Dívida quitada, saldo em caixa e reserva são fatos verificáveis em balanço, com certidão e ata — sem discussão futura sobre quem pode retirar o quê."]]),

    // ---------- metas ----------
    H("As três metas — cumulativas"),
    new Table({columnWidths:W,rows:[
      metaRow("META I","Extinção do perímetro","100% das dívidas","Originadas nos negócios conduzidos pelo patriarca — em nome da Grison e Cia, das pessoas físicas ou das satélites, seja qual for o nome no contrato."),
      metaRow("META II","Caixa-pulmão","R$ 2.000.000,00","Caixa livre consolidado, sustentado por dois balancetes consecutivos. Impede que a próxima parcela ou multa vire nova dívida."),
      metaRow("META III","Moradia da usufrutuária","R$ 1.060.000,00 *","Reserva vinculada para o imóvel residencial da Sra. Elman, ou imóvel já adquirido. Vem antes de qualquer retirada dos filhos."),
    ]}),
    R([["Fora da conta: ",{bold:true,size:19}],["o adiantamento do loteador parceiro (Santa Helena) não entra na Meta I. Ele é autoliquidável — o próprio loteador retém os valores devidos ao grupo até recompor o adiantamento e só depois repassa o excedente. Não há o que quitar, há o que aguardar.",{size:19,color:MUT}]],{spacing:sp(120,0)}),

    // ---------- minuta ----------
    H("Minuta — Cláusula [•]: destinação de resultados e metas de liberação"),

    R([["1. Princípio. ",{bold:true}],["A distribuição de lucros, dividendos, juros sobre capital próprio ou qualquer outra forma de retorno do capital aos nu-proprietários não é regida por prazo, mas pelo cumprimento cumulativo das três Metas do item 3, tenha esse cumprimento a duração que tiver. Até lá, os resultados são integralmente retidos e destinados a esse cumprimento."]]),

    R([["2. Definições.",{bold:true}]],{spacing:sp(60,60)}),
    R([["2.1  Perímetro de Endividamento Familiar. ",{bold:true}],["Todas as obrigações pecuniárias, de qualquer natureza e ainda que vincendas, originadas nos negócios conduzidos, liderados ou orientados pelo Sr. Wilson Grison, "],["independentemente da pessoa física ou jurídica em cujo nome tenham sido contraídas",{bold:true}],[": as da Grison e Cia Ltda.; as pessoais de Wilson e de Elman M. Coelho Grison como devedores, avalistas, fiadores ou coobrigados; e as das sociedades por eles constituídas, controladas ou dirigidas de fato, ainda que registradas em nome de filhos, cônjuge ou terceiros — EMC Grison, Bioaçaí, ML Serviços Agrícolas, Natyrê, Tiúba 2 SPE e Patrimônio Digital. A relação consta do Anexo I, aprovado em até 90 dias e atualizado semestralmente; obrigação não listada que se enquadre nesta definição integra o Perímetro."]],{indent:{left:200}}),
    R([["2.2  Dívida Quitada. ",{bold:true}],["Aquela cuja exigibilidade se extinguiu em face de todos os coobrigados — por pagamento, deságio, transação, novação, decisão transitada em julgado, prescrição ou assunção liberatória com exoneração dos garantidores. É indiferente o valor desembolsado: exige-se que a obrigação deixe de ser exigível do Perímetro."]],{indent:{left:200}}),
    R([["2.3  Dívida Equacionada. ",{bold:true}],["A obrigação em parcelamento ou transação regularmente deferida que, cumulativamente, esteja adimplente, não conte com garantia pessoal de qualquer dos nu-proprietários e tenha o saldo devedor integralmente provisionado em reserva vinculada não computável na Meta II. Vale como cumprida enquanto mantidos os requisitos."]],{indent:{left:200}}),
    R([["2.4  Adiantamento Santa Helena. ",{bold:true}],["Valores antecipados pelo loteador parceiro por conta de resultados futuros da parceria, recompostos por retenção, pelo próprio loteador, dos repasses devidos ao grupo. "],["Está expressamente excluído do Perímetro",{bold:true}],[" e não conta para a Meta I, por constituir obrigação de liquidação automática por compensação contratual, independente de desembolso ou de ato de gestão dos sócios. Consta do Anexo I apenas para informação."]],{indent:{left:200}}),
    R([["2.5  Caixa Mínimo. ",{bold:true}],["Caixa e equivalentes de livre movimentação da Companhia e controladas, excluídos os recursos vinculados à Meta III, os provisionados no item 2.3, os de terceiros e os valores bloqueados ou de destinação vinculada."]],{indent:{left:200}}),

    R([["3. As Metas — cumulativas e não fracionáveis. ",{bold:true}],["(I) 100% das obrigações do Perímetro Quitadas ou Equacionadas, ressalvado o Adiantamento Santa Helena; (II) Caixa Mínimo igual ou superior a R$ 2.000.000,00, mantido por dois balancetes mensais consecutivos; (III) Reserva Moradia igual ou superior a R$ 1.060.000,00 destinada ao imóvel residencial da Sra. Elman, ou imóvel já adquirido e quitado. Os valores das Metas II e III são corrigidos anualmente pelo IPCA. O cumprimento parcial, ainda que de duas Metas, não autoriza distribuição parcial, proporcional ou antecipada."]],{spacing:sp(60,80)}),

    R([["4. Retenção. ",{bold:true}],["Até a Data de Liberação é vedada aos nu-proprietários qualquer distribuição de lucros, dividendos, juros sobre capital próprio, redução de capital com restituição, resgate ou reembolso, ressalvado o item 5. Os lucros retidos vão para reserva fundamentada em orçamento de capital aprovado anualmente em Assembleia (art. 196 da Lei 6.404/76), e o Estatuto fixará o dividendo obrigatório em percentual compatível com esta Cláusula (art. 202), sob pena de conflito entre os dois documentos."]]),

    R([["5. Exceções. ",{bold:true}],["Não configuram distribuição vedada: (a) a "],["retirada mensal dos usufrutuários",{bold:true}],[" Wilson e Elman, de até R$ [•] cada, reajustável pelo IPCA, a título dos frutos que lhes pertencem por direito — a trava recai sobre a retirada dos filhos, não sobre o usufruto reservado na doação; (b) as parcelas do ITCMD da doação e os custos dos atos societários e registrais correlatos; (c) o pró-labore, em valor de mercado e aprovado em Assembleia, dos sócios que exerçam função; (d) o reembolso de despesas comprovadas no interesse da Companhia; (e) a devolução de mútuo formalizado. Os valores de (a) e (c) constam do Anexo II e sua majoração não pode ter por efeito substituir a distribuição vedada."]]),

    R([["6. Vedação à distribuição disfarçada. ",{bold:true}],["Durante a retenção é vedado o pagamento de despesas pessoais pela Companhia ou controladas, o mútuo a sócio ou parte relacionada sem contrato e juros de mercado, a contratação de parte relacionada fora de mercado, o uso de bem social sem contrapartida e o desvio de oportunidade de negócio para veículo pessoal de sócio. As sociedades do grupo manterão contabilidades segregadas e contratos escritos entre si. A violação obriga à restituição integral e corrigida."]]),

    R([["7. Verificação. ",{bold:true}],["A administração apresenta Relatório de Progresso das Metas semestralmente. Entendendo-as cumpridas, convoca Assembleia instruída com balancete de até 30 dias, declaração do contador, comprovação da extinção de cada obrigação e certidões das pessoas físicas e jurídicas envolvidas, que declara em ata a Data de Liberação. Qualquer sócio pode requerer, às suas expensas, verificação independente antes da deliberação."]]),

    R([["8. Efeitos e restabelecimento da trava. ",{bold:true}],["Aberta a distribuição, ela não pode reduzir o Caixa Mínimo abaixo do valor da Meta II — que deixa de ser meta e passa a ser piso permanente — e a Reserva Moradia segue indisponível até a aquisição. A distribuição fica automaticamente suspensa se for revelada obrigação do Perímetro anterior à declaração e dela não constante, se a Dívida Equacionada deixar de atender ao item 2.3, ou se o caixa cair abaixo do piso."]]),

    R([["9. Ausência de prazo. ",{bold:true}],["Não há prazo mínimo nem máximo, e nenhum decurso de tempo produz liberação automática. Decorridos 60 meses sem cumprimento integral, a Assembleia revisa o plano de execução e o cronograma; a revisão não flexibiliza nem dispensa as Metas, o que dependerá de deliberação específica pelo quórum de alteração do Acordo."]]),

    R([["Anexos. ",{bold:true}],["I — Mapa de Dívidas do Perímetro (credor, natureza, situação processual, garantias, coobrigados e status), com o Adiantamento Santa Helena destacado como item excluído. II — Retirada dos usufrutuários e pró-labore aprovados.",{}]]),

    // ---------- pendências ----------
    H("Três definições pendentes antes de fechar a redação"),
    R([["1.  Valor do imóvel da Meta III. ",{bold:true}],["A minuta está com R$ 1.060.000,00. Confirmar se é este o número ou R$ 1.600.000,00."]],{indent:{left:200}}),
    R([["2.  Retirada dos usufrutuários (item 5, alínea a). ",{bold:true}],["Definir o valor mensal que Wilson e Elman seguem recebendo durante a retenção. A cláusula não pode ser um “ninguém recebe nada” puro: o usufruto é dos pais e os frutos lhes pertencem por direito."]],{indent:{left:200}}),
    R([["3.  Dívidas em parcelamento oficial (item 2.3). ",{bold:true}],["Decidir se contam como cumpridas quando adimplentes e integralmente cobertas por reserva. Sem essa regra, uma transação da PGFN em 120 meses trava a distribuição por dez anos sozinha."]],{indent:{left:200}}),

  ]}]
});

Packer.toBuffer(doc).then(b=>{fs.writeFileSync("/home/user/leogrison/docs/Minuta-Distribuicao-por-Metas.docx",b);console.log("ok",b.length)});
