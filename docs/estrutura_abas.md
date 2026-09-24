# Estrutura recomendada do arquivo Excel

## Dados_Colaboradores

A aba deve conter a base original e uma tabela nomeada, por exemplo `tblColaboradores`. As colunas devem ser padronizadas antes da construção das análises. O ID deve ser único, os salários devem ser numéricos, os departamentos devem ter uma nomenclatura única e as datas de admissão devem seguir o mesmo formato.

## Auxiliares

Esta aba pode concentrar listas de apoio, médias por cargo e regras utilizadas para preencher salários ausentes ou zerados. Manter essas regras separadas facilita a auditoria e reduz fórmulas espalhadas pelo relatório.

## Tabelas_Dinamicas

As tabelas dinâmicas devem ser construídas sobre `tblColaboradores`. Recomenda-se manter uma tabela para cada análise: headcount por departamento, folha por departamento, salário por nível, salário por região e satisfação média por departamento.

## Dashboard

A aba executiva deve apresentar os KPIs na parte superior, os gráficos no centro e as segmentações em uma área fixa. O botão de atualização deve ficar visível e utilizar a macro documentada em `vba/limpar_filtros_atualizar.bas`.
