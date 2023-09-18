select
	case
		when tnf.nfse = 1 then 'NFSe'
		else 'NF'
	end as 'Tipo',
	tnf.numero as 'NF',
	convert(date,tnf.dt_emissao) as 'Data Emissão',
	tnf.cancelada as 'Cancelada',
	case
		when tnf.cancelada = 1 then 'Cancelada'
		when tnf.numero in (select distinct nfref_numero_nf from tnota_fiscal_nfref where cod_empresa =1) then 'Devolvida pela NF ' + convert(char,(select top 1 numero_nf from tnota_fiscal_nfref where cod_empresa = 1 and nfref_numero_nf = tnf.numero))
		when tnf.nfse = 1 then 'Autorizada'
		when tnf.nfe_desc_status is null or tnf.nfe_desc_status = '' then 'Aguardando Aprovação'
		else 'Autorizada'
	end as 'Status',
	tnf.cfop as 'CFOP',
	tnf.natureza_op as 'Natureza Operação',
	tnf.entrada as 'Entrada',
	tnf.saida as 'Saída',
	tnfi.codigo as 'Cod Interno',
	isnull(tos.n_os,(select top 1 n_os from tos where cod_empresa=1 and n_os not like 'e%'and cod_pecas=tnfi.codigo)) as 'OS',
	tnf.razao_social as 'Cliente',
	isnull(tnfi.produto,'SERVIÇO') as 'Produto',
	isnull(tnfi.qtde,tnfse.qtde) as 'Qtd',
	isnull(tnfi.unidade,tnfse.unidade) as 'Un',
	isnull(tnfi.vl_total+tnfi.vl_ipi,tnfse.vl_total) as 'Valor',
	(select top 1 codigo from tordem_fat where cod_empresa = 1 and status=1 and numero_nf = tnf.numero) as 'OF'
from
	(select * from tnota_fiscal_item where cod_empresa=1) as tnfi

right join (select * from tnota_fiscal where cod_empresa=1 and condicao_pagamento is not null and condicao_pagamento <> '' and cli_for=0) as tnf
	on tnfi.numero_nf = tnf.numero
		left join (select * from tnota_fiscal_servico where cod_empresa = 1) as tnfse
			on tnf.numero = tnfse.numero_nf

left join (select * from torcamento_itens where cod_empresa=1) as toi
	on tnfi.guid_orcamento = toi.guid_linha
		left join (select * from torcamento where cod_empresa = 1 and estagio_orc = 'GANHOU' and status=1) as tor
			on toi.cod_orcamento = tor.codigo

left join (select * from tos where cod_empresa=1 and n_os not like 'e%') as tos
	on tnfi.guid_orcamento = tos.guid_orcamento

--order by tnf.dt_emissao desc, tnf.numero desc