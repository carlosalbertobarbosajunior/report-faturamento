# report-faturamento
Script para gerar uma tabela em HTML, disparada diáriamente por e-mail, para alertar a direção sobre o valor de faturamento atual.

## Problemática

A alta direção necessita averiguar - ao menos diáriamente - as informações de faturamento da empresa. Os dados precisam conter as informações de notas fiscais de movimentações de saída, apenas para os CFOP's destinados à venda de projetos.

Para isso, foi necessário desenvolver um script capaz de buscar as informações de emissão de NFs no banco de dados CPS e tratá-las para resultar em um relatório de faturamento do mês. A metodologia escolhida para envio dessas informações foi envio de e-mails, todos os dias pela manhã.


## Arquivo Executável

Para facilitar a emissão do relatório em qualquer ambiente, há um arquivo executável desenvolvido através da biblioteca **pyinstaller**. O arquivo .exe se encontra em *dist\report-faturamento\report-faturamento.exe*.


## Funções e desenvolvimento do script

O desenvolvimento do script pode ser visto no arquivo **notebook.ipynb** e contempla as bibliotecas e as funções utilizadas em sua totalidade. Há uma versão com extensão .py nomeada de **report-faturamento.py**.

### is_first_business_day():
    '''
    Output:
      today == first_day       - Boolean que identifica se hoje é
                                 o primeiro dia útil do mês
    
    Ação:
      Calcula o primeiro dia útil do mês vigente e o compara com o
      dia de hoje.
    '''

### create_df_from_database(config_file, query):
    '''
    Inputs:
      config_file - Caminho do arquivo .ini (string) que possui as 
                    configurações de conexão com o banco de dados 
                    da empresa.
      query       - Instrução em SQL (string) que será transformada
                    em um dataframe do pandas.
                    
    Outputs:
      df          - Dataframe do pandas construído a partir da instrução
                    SQL da variável query.
    '''

### format_cell(valor):
    '''
    Input:
      valor          - células do dataframe
      
    Output:
      Múltiplos. Aplica formatação condicional para cada célula
      baseada no valor da mesma.
    '''


Para acessar as informações no banco de dados CPS, foi desenvolvida uma view chamada vwNFsProjetos. O DER e seu script estão descritos nas próximas seções.

## Diagrama de Entidade Relacional

<img src="DER vwNFsProjetos.png">

## Consulta ao banco de dados

O script está disponibilizado no arquivo **vwNFsProjetos.sql**. 

  ```sql
select
	-- Classificador de NF e NFSe
	case
		when tnf.nfse = 1 then 'NFSe'
		else 'NF'
	end as 'Tipo',
	tnf.numero as 'NF',
	convert(date,tnf.dt_emissao) as 'Data Emissão',
	tnf.cancelada as 'Cancelada',
	case
		-- Condicional de NF cancelada
		when tnf.cancelada = 1 then 'Cancelada'
		-- Condicional de NF devolvida
		when tnf.numero in (select distinct nfref_numero_nf from tnota_fiscal_nfref where cod_empresa =1) then 'Devolvida pela NF ' + convert(char,(select top 1 numero_nf from tnota_fiscal_nfref where cod_empresa = 1 and nfref_numero_nf = tnf.numero))
		-- NFSe's importadas da prefeitura (autorizadas sempre)
		when tnf.nfse = 1 then 'Autorizada'
		-- NFs sem status no GRV ainda estão aguardando aprovação
		when tnf.nfe_desc_status is null or tnf.nfe_desc_status = '' then 'Aguardando Aprovação'
		-- Caso contrário, autorizada
		else 'Autorizada'
	end as 'Status',
	tnf.cfop as 'CFOP',
	tnf.natureza_op as 'Natureza Operação',
	tnf.entrada as 'Entrada',
	tnf.saida as 'Saída',
	tnfi.codigo as 'Cod Interno',
	-- Caso não tenha ordem de serviço, tenta buscar a mesma pelo código de cadastro do material
	isnull(tos.n_os,(select top 1 n_os from tos where cod_empresa=1 and n_os not like 'e%'and cod_pecas=tnfi.codigo)) as 'OS',
	tnf.razao_social as 'Cliente',
	-- Caso o produto não tenha nome, chama-o de serviço
	isnull(tnfi.produto,'SERVIÇO') as 'Produto',
	isnull(tnfi.qtde,tnfse.qtde) as 'Qtd',
	isnull(tnfi.unidade,tnfse.unidade) as 'Un',
	isnull(tnfi.vl_total+tnfi.vl_ipi,tnfse.vl_total) as 'Valor',
	-- Identifica a ordem de faturamento que gerou a NF, caso haja
	(select top 1 codigo from tordem_fat where cod_empresa = 1 and status=1 and numero_nf = tnf.numero) as 'OF'
from
	(select * from tnota_fiscal_item where cod_empresa=1) as tnfi

-- Relacionamento com a tabela principal de notas fiscais
right join (select * from tnota_fiscal where cod_empresa=1 and condicao_pagamento is not null and condicao_pagamento <> '' and cli_for=0) as tnf
	on tnfi.numero_nf = tnf.numero
		-- Relacionamento com a tabela de notas fiscais de serviço
		left join (select * from tnota_fiscal_servico where cod_empresa = 1) as tnfse
			on tnf.numero = tnfse.numero_nf

-- Relacionamento com as tabelas de orçamentos
left join (select * from torcamento_itens where cod_empresa=1) as toi
	on tnfi.guid_orcamento = toi.guid_linha
		left join (select * from torcamento where cod_empresa = 1 and estagio_orc = 'GANHOU' and status=1) as tor
			on toi.cod_orcamento = tor.codigo

-- Relacionamento com a tabela de ordens de serviço
left join (select * from tos where cod_empresa=1 and n_os not like 'e%') as tos
	on tnfi.guid_orcamento = tos.guid_orcamento

--order by tnf.dt_emissao desc, tnf.numero desc
  ```


