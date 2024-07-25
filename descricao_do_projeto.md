# Robô de Monitoramento Diário de Preço

 -> Foco do projeto: Automação de Processos
    Nível do projeto: Intermediário


## Breve descrição do projeto:
  Crie um script em Python que automatize a consulta de preços de um único produto(ex: iphone
  15 pro max,você escolhe o produto) em algum site da sua escolha(você pode escolher
  qualquer site mesmo) e atualize uma planilha Excel com os preços coletados de 30 em 30
  minutos
## Funcionalidades que o projeto deve possuir:
    1. Consulta Automatizada:
      ○ Acesse um site que venda o produto que escolheu.
      ○ Verificar o preço atual.
      ○ Guardar o valor do preço(somente o valor numérico, não em texto)
        i. ex:
        ii. Se o valor está como R$1500,00 no site, você irá guardar apenas 1500
        iii. Se o valor está como R$1700,50 no site, você irá guardar apenas
        1700.50
    2. Manipulação de Planilhas:
      ○ Crie uma planilha com a seguintes colunas:
        i. Produto(que armazena o nome do produto)
        ii. Data atual(que corresponde à data da consulta)
        iii. Valor
        iv. Link(link direto para o produto)
    3. Automatização Recorrente:
      ○ Criar um agendamento para que o bot rode de 30 em 30 minutos
