# Cascalheira — análise de vendas

Análise do histórico de vendas da cascalheira do grupo, feita para dar base à decisão sobre a
parceria de extração na área do vizinho.

## Arquivos

| Arquivo | O que é |
|---|---|
| `analise-vendas.py` | Lê o export do sistema e consolida os números. É onde moram os critérios de classificação. |
| `atualizar-relatorio.py` | Roda a análise e injeta os números no relatório visual. |
| `relatorio-vendas.html` | O relatório publicado. Texto e desenho à mão; números vindos do bloco `<script id="dados">`. |

O export do sistema **não fica no repositório** — traz nome, CPF e CNPJ de 289 clientes. Guarde o
`.xls` fora do controle de versão e passe o caminho por argumento.

## Atualizar com um export novo

```bash
pip install xlrd
python3 atualizar-relatorio.py /caminho/para/relatorio.xls
```

O script reescreve só o bloco de dados dentro do HTML. Depois é só republicar o Artifact no mesmo
endereço.

## Critérios que a análise aplica

Nenhum deles vem de um campo de cadastro — todos são inferidos, e os três primeiros devem ser
conferidos contra a operação real antes de a análise virar decisão.

1. **Frações fora.** Onze linhas `Fração NN (m²)` são venda de lotes do Serra Bella, não cascalho.
   Removidas integralmente (R$ 1.955.015,65).
2. **Retirada × entrega pela faixa de preço.** Cada produto tem dois preços bem separados — pátio e
   entregue. O corte é achado no maior vão da distribuição de preços de cada produto, não num
   limiar fixo: o Aterro 12m³ separa em R$ 249, o Sub Base 12m³ em R$ 325.
3. **Frete repassado.** O que o preço entregue excede o preço de pátio do mesmo produto é frete, e
   sai inteiro para o caminhoneiro. Descontado do faturamento para chegar à receita própria.
4. **Correções de cadastro.** Uma linha de "Aterro 12m³" da Prefeitura a R$ 14,50 é preço por m³, e
   não por carga — tratada como granel. Uma linha da Construtora Alja a R$ 19.040 em uma única
   unidade é contrato fechado, sem volume real; fica num canal próprio e fora do cálculo de frete.

## O que ainda falta

- **Export com data de emissão.** O relatório atual não tem coluna de data, então não há série
  mensal nem leitura de sazonalidade — que em cascalho é o que manda, porque a chuva para a
  extração.
- **Custos da operação.** Máquina, combustível, mão de obra, manutenção e administração, para
  fechar a margem por m³.
- **Números da área nova.** Base de cálculo do royalty, volume da jazida, distâncias, prazo, custo
  de licenciamento na ANM e o passivo de recuperação.

## Pedido ao financeiro

`pedido-financeiro.html` é a lista de levantamento a ser enviada ao financeiro: o relatório que o
Bling precisa exportar (Vendas › Pedidos de venda › Exportar, com data por linha) e as perguntas
de custo operacional. Escopo deliberadamente limitado a opex — licenciamento, recuperação de área
e capex ficam com o advogado e com o especialista em direito minerário.

A lista dos 45 clientes classificados como revenda é gerada sob demanda a partir do export e não
fica versionada, pelo mesmo motivo do `.xls`.
