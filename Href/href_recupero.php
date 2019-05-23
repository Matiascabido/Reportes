<?php
  include 'conn.php';
  $mes = $_POST['fecham'];
  $año = $_POST['fechaa'];

	$sql = "select 'Cartera Cumplimiento Normal' as Periodo, count(*) as Cantidad,  cast(sum([Calculo_Mora]) as decimal(19,2)) as Monto,
	cast(sum(Total_Pagado) as decimal(19,2)) as Pagado, cast(sum([Saldo_Deudor_Calculado]) as decimal(19,2)) as SALDO_DEUDOR
	from
	(
	SELECT 
	cast((C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * isnull((SELECT count(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS NULL and CU.cuo_monto_capital > 0 and CU.cuo_vencimiento_1 <=  EOMONTH(DATEFROMPARTS(".$año.",".$mes.",1))),0) as decimal(19,2)) as [Calculo_Mora],
	cast((C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * (select count(*) from fbc_cuota cccc  inner join fbc_pago p on p.pag_id = cccc.id_pago where cccc.id_credito = c.cre_id and cccc.id_pago is not null and p.pag_fecha_pago <= EOMONTH(datefromparts(".$año.",".$mes.",1))) as decimal(19,2)) as [Total_Pagado],
	cast(C.CRE_MONTO_OTORGADO - (C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * (select count(*) from fbc_cuota cccc  inner join fbc_pago p on p.pag_id = cccc.id_pago where cccc.id_credito = c.cre_id and cccc.id_pago is not null and cccc.cuo_monto_capital > 0 and p.pag_fecha_pago <= EOMONTH(datefromparts(".$año.",".$mes.",1))) as decimal(19,2)) as [Saldo_Deudor_Calculado]
	FROM FBC_CREDITO C 
	INNER JOIN FBC_SOLICITUD S ON S.SOL_ID = C.ID_SOLICITUD 
	INNER JOIN FBC_PERSONA P ON P.PER_ID = S.ID_TITULAR 
	INNER JOIN FBC_LOCALIDAD L ON L.LOC_ID = P.ID_LOCALIDAD 
	inner join fbc_programa pr on pr.pro_id = s.id_programa_solicitado 
	WHERE 1 = 1 AND c.cre_fecha_efectivizacion  <= EOMONTH(datefromparts(".$año.",".$mes.",1))
	and S.id_programa_solicitado = 43108
	and (C.CRE_MONTO_OTORGADO - isnull(((SELECT SUM(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS not NULL)),0)) > 0
	and isnull((SELECT count(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS NULL  and CU.cuo_monto_capital > 0 and CU.cuo_vencimiento_1 <=  EOMONTH(DATEFROMPARTS(".$año.",".$mes.",1))),0) = 0
	) as X
	union all
	select '1-29 dias.' as Periodo, count(*) as Cantidad,  cast(sum([Calculo_Mora]) as decimal(19,2)) as Monto,
	cast(sum(Total_Pagado) as decimal(19,2)) as Pagado, cast(sum([Saldo_Deudor_Calculado]) as decimal(19,2)) as SALDO_DEUDOR
	from
	(
	SELECT 
	cast((C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * isnull((SELECT count(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS NULL and CU.cuo_monto_capital > 0 and CU.cuo_vencimiento_1 <=  EOMONTH(DATEFROMPARTS(".$año.",".$mes.",1))),0) as decimal(19,2)) as [Calculo_Mora],
	cast((C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * (select count(*) from fbc_cuota cccc  inner join fbc_pago p on p.pag_id = cccc.id_pago where cccc.id_credito = c.cre_id and cccc.id_pago is not null and p.pag_fecha_pago <= EOMONTH(datefromparts(".$año.",".$mes.",1))) as decimal(19,2)) as [Total_Pagado],
	cast(C.CRE_MONTO_OTORGADO - (C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * (select count(*) from fbc_cuota cccc  inner join fbc_pago p on p.pag_id = cccc.id_pago where cccc.id_credito = c.cre_id and cccc.id_pago is not null and cccc.cuo_monto_capital > 0 and p.pag_fecha_pago <= EOMONTH(datefromparts(".$año.",".$mes.",1))) as decimal(19,2)) as [Saldo_Deudor_Calculado]
	FROM FBC_CREDITO C 
	INNER JOIN FBC_SOLICITUD S ON S.SOL_ID = C.ID_SOLICITUD 
	INNER JOIN FBC_PERSONA P ON P.PER_ID = S.ID_TITULAR 
	INNER JOIN FBC_LOCALIDAD L ON L.LOC_ID = P.ID_LOCALIDAD 
	inner join fbc_programa pr on pr.pro_id = s.id_programa_solicitado 
	WHERE 1 = 1 AND c.cre_fecha_efectivizacion  <= EOMONTH(datefromparts(".$año.",".$mes.",1))
	and S.id_programa_solicitado = 43108
	and (C.CRE_MONTO_OTORGADO - isnull(((SELECT SUM(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS not NULL)),0)) > 0
	and isnull((SELECT count(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS NULL  and CU.cuo_monto_capital > 0 and CU.cuo_vencimiento_1 <=  EOMONTH(DATEFROMPARTS(".$año.",".$mes.",1))),0) = 1
	) as X
	union all
	select '30-90 dias.' as Periodo, count(*) as Cantidad,  cast(sum([Calculo_Mora]) as decimal(19,2)) as Monto,
	cast(sum(Total_Pagado) as decimal(19,2)) as Pagado, cast(sum([Saldo_Deudor_Calculado]) as decimal(19,2)) as SALDO_DEUDOR
	from
	(
	SELECT 
	cast((C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * isnull((SELECT count(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS NULL and CU.cuo_monto_capital > 0 and CU.cuo_vencimiento_1 <=  EOMONTH(DATEFROMPARTS(".$año.",".$mes.",1))),0) as decimal(19,2)) as [Calculo_Mora],
	cast((C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * (select count(*) from fbc_cuota cccc  inner join fbc_pago p on p.pag_id = cccc.id_pago where cccc.id_credito = c.cre_id and cccc.id_pago is not null and p.pag_fecha_pago <= EOMONTH(datefromparts(".$año.",".$mes.",1))) as decimal(19,2)) as [Total_Pagado],
	cast(C.CRE_MONTO_OTORGADO - (C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * (select count(*) from fbc_cuota cccc  inner join fbc_pago p on p.pag_id = cccc.id_pago where cccc.id_credito = c.cre_id and cccc.id_pago is not null and cccc.cuo_monto_capital > 0 and p.pag_fecha_pago <= EOMONTH(datefromparts(".$año.",".$mes.",1))) as decimal(19,2)) as [Saldo_Deudor_Calculado]
	FROM FBC_CREDITO C 
	INNER JOIN FBC_SOLICITUD S ON S.SOL_ID = C.ID_SOLICITUD 
	INNER JOIN FBC_PERSONA P ON P.PER_ID = S.ID_TITULAR 
	INNER JOIN FBC_LOCALIDAD L ON L.LOC_ID = P.ID_LOCALIDAD 
	inner join fbc_programa pr on pr.pro_id = s.id_programa_solicitado 
	WHERE 1 = 1 AND c.cre_fecha_efectivizacion  <= EOMONTH(datefromparts(".$año.",".$mes.",1))
	and S.id_programa_solicitado = 43108
	and (C.CRE_MONTO_OTORGADO - isnull(((SELECT SUM(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS not NULL)),0)) > 0
	and isnull((SELECT count(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS NULL  and CU.cuo_monto_capital > 0 and CU.cuo_vencimiento_1 <=  EOMONTH(DATEFROMPARTS(".$año.",".$mes.",1))),0) in (2,3)
	) as X
	union all
	select '91 - 180 dias.' as Periodo, count(*) as Cantidad,  cast(sum([Calculo_Mora]) as decimal(19,2)) as Monto,
	cast(sum(Total_Pagado) as decimal(19,2)) as Pagado, cast(sum([Saldo_Deudor_Calculado]) as decimal(19,2)) as SALDO_DEUDOR
	from
	(
	SELECT 
	cast((C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * isnull((SELECT count(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS NULL and CU.cuo_monto_capital > 0 and CU.cuo_vencimiento_1 <=  EOMONTH(DATEFROMPARTS(".$año.",".$mes.",1))),0) as decimal(19,2)) as [Calculo_Mora],
	cast((C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * (select count(*) from fbc_cuota cccc  inner join fbc_pago p on p.pag_id = cccc.id_pago where cccc.id_credito = c.cre_id and cccc.id_pago is not null and p.pag_fecha_pago <= EOMONTH(datefromparts(".$año.",".$mes.",1))) as decimal(19,2)) as [Total_Pagado],
	cast(C.CRE_MONTO_OTORGADO - (C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * (select count(*) from fbc_cuota cccc  inner join fbc_pago p on p.pag_id = cccc.id_pago where cccc.id_credito = c.cre_id and cccc.id_pago is not null and cccc.cuo_monto_capital > 0 and p.pag_fecha_pago <= EOMONTH(datefromparts(".$año.",".$mes.",1))) as decimal(19,2)) as [Saldo_Deudor_Calculado]
	FROM FBC_CREDITO C 
	INNER JOIN FBC_SOLICITUD S ON S.SOL_ID = C.ID_SOLICITUD 
	INNER JOIN FBC_PERSONA P ON P.PER_ID = S.ID_TITULAR 
	INNER JOIN FBC_LOCALIDAD L ON L.LOC_ID = P.ID_LOCALIDAD 
	inner join fbc_programa pr on pr.pro_id = s.id_programa_solicitado 
	WHERE 1 = 1 AND c.cre_fecha_efectivizacion  <= EOMONTH(datefromparts(".$año.",".$mes.",1))
	and S.id_programa_solicitado = 43108
	and (C.CRE_MONTO_OTORGADO - isnull(((SELECT SUM(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS not NULL)),0)) > 0
	and isnull((SELECT count(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS NULL  and CU.cuo_monto_capital > 0 and CU.cuo_vencimiento_1 <=  EOMONTH(DATEFROMPARTS(".$año.",".$mes.",1))),0) in (4,5,6)
	) as X
	union all
	select '> 180 dias.' as Periodo, count(*) as Cantidad,  cast(sum([Calculo_Mora]) as decimal(19,2)) as Monto,
	cast(sum(Total_Pagado) as decimal(19,2)) as Pagado, cast(sum([Saldo_Deudor_Calculado]) as decimal(19,2)) as SALDO_DEUDOR
	from
	(
	SELECT 
	cast((C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * isnull((SELECT count(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS NULL and CU.cuo_monto_capital > 0 and CU.cuo_vencimiento_1 <=  EOMONTH(DATEFROMPARTS(".$año.",".$mes.",1))),0) as decimal(19,2)) as [Calculo_Mora],
	cast((C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * (select count(*) from fbc_cuota cccc  inner join fbc_pago p on p.pag_id = cccc.id_pago where cccc.id_credito = c.cre_id and cccc.id_pago is not null and p.pag_fecha_pago <= EOMONTH(datefromparts(".$año.",".$mes.",1))) as decimal(19,2)) as [Total_Pagado],
	cast(C.CRE_MONTO_OTORGADO - (C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * (select count(*) from fbc_cuota cccc  inner join fbc_pago p on p.pag_id = cccc.id_pago where cccc.id_credito = c.cre_id and cccc.id_pago is not null and cccc.cuo_monto_capital > 0 and p.pag_fecha_pago <= EOMONTH(datefromparts(".$año.",".$mes.",1))) as decimal(19,2)) as [Saldo_Deudor_Calculado]
	FROM FBC_CREDITO C 
	INNER JOIN FBC_SOLICITUD S ON S.SOL_ID = C.ID_SOLICITUD 
	INNER JOIN FBC_PERSONA P ON P.PER_ID = S.ID_TITULAR 
	INNER JOIN FBC_LOCALIDAD L ON L.LOC_ID = P.ID_LOCALIDAD 
	inner join fbc_programa pr on pr.pro_id = s.id_programa_solicitado 
	WHERE 1 = 1 AND c.cre_fecha_efectivizacion  <= EOMONTH(datefromparts(".$año.",".$mes.",1))
	and S.id_programa_solicitado = 43108
	and (C.CRE_MONTO_OTORGADO - isnull(((SELECT SUM(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS not NULL)),0)) > 0
	and isnull((SELECT count(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS NULL  and CU.cuo_monto_capital > 0 and CU.cuo_vencimiento_1 <=  EOMONTH(DATEFROMPARTS(".$año.",".$mes.",1))),0) >= 7
	) as X"; 

	$sqlNo = "select 'Cartera Cumplimiento Normal' as Periodo, count(*) as Cantidad,  cast(sum([Calculo_Mora]) as decimal(19,2)) as Monto,
	cast(sum(Total_Pagado) as decimal(19,2)) as Pagado, cast(sum([Saldo_Deudor_Calculado]) as decimal(19,2)) as SALDO_DEUDOR
	from
	(
	SELECT 
	cast((C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * isnull((SELECT count(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS NULL and CU.cuo_monto_capital > 0 and CU.cuo_vencimiento_1 <=  EOMONTH(DATEFROMPARTS(".$año.",".$mes.",1))),0) as decimal(19,2)) as [Calculo_Mora],
	cast((C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * (select count(*) from fbc_cuota cccc  inner join fbc_pago p on p.pag_id = cccc.id_pago where cccc.id_credito = c.cre_id and cccc.id_pago is not null and p.pag_fecha_pago <= EOMONTH(datefromparts(".$año.",".$mes.",1))) as decimal(19,2)) as [Total_Pagado],
	cast(C.CRE_MONTO_OTORGADO - (C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * (select count(*) from fbc_cuota cccc  inner join fbc_pago p on p.pag_id = cccc.id_pago where cccc.id_credito = c.cre_id and cccc.id_pago is not null and cccc.cuo_monto_capital > 0 and p.pag_fecha_pago <= EOMONTH(datefromparts(".$año.",".$mes.",1))) as decimal(19,2)) as [Saldo_Deudor_Calculado]
	FROM FBC_CREDITO C 
	INNER JOIN FBC_SOLICITUD S ON S.SOL_ID = C.ID_SOLICITUD 
	INNER JOIN FBC_PERSONA P ON P.PER_ID = S.ID_TITULAR 
	INNER JOIN FBC_LOCALIDAD L ON L.LOC_ID = P.ID_LOCALIDAD 
	inner join fbc_programa pr on pr.pro_id = s.id_programa_solicitado 
	WHERE 1 = 1 AND c.cre_fecha_efectivizacion  <= EOMONTH(datefromparts(".$año.",".$mes.",1))
	and S.id_programa_solicitado <> 43108
	and (C.CRE_MONTO_OTORGADO - isnull(((SELECT SUM(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS not NULL)),0)) > 0
	and isnull((SELECT count(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS NULL  and CU.cuo_monto_capital > 0 and CU.cuo_vencimiento_1 <=  EOMONTH(DATEFROMPARTS(".$año.",".$mes.",1))),0) = 0
	and s.sol_id not in (select id_solicitud from FBC_SEGUIMIENTO_SOLICITUD where sso_estado in ('CIERRE_ANTICIPADO','CIERRE_EXTRAORDINARIO','CERRADO'))
	) as X
	union all
	select '1-29 dias.' as Periodo, count(*) as Cantidad,  cast(sum([Calculo_Mora]) as decimal(19,2)) as Monto,
	cast(sum(Total_Pagado) as decimal(19,2)) as Pagado, cast(sum([Saldo_Deudor_Calculado]) as decimal(19,2)) as SALDO_DEUDOR
	from
	(
	SELECT 
	cast((C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * isnull((SELECT count(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS NULL and CU.cuo_monto_capital > 0 and CU.cuo_vencimiento_1 <=  EOMONTH(DATEFROMPARTS(".$año.",".$mes.",1))),0) as decimal(19,2)) as [Calculo_Mora],
	cast((C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * (select count(*) from fbc_cuota cccc  inner join fbc_pago p on p.pag_id = cccc.id_pago where cccc.id_credito = c.cre_id and cccc.id_pago is not null and p.pag_fecha_pago <= EOMONTH(datefromparts(".$año.",".$mes.",1))) as decimal(19,2)) as [Total_Pagado],
	cast(C.CRE_MONTO_OTORGADO - (C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * (select count(*) from fbc_cuota cccc  inner join fbc_pago p on p.pag_id = cccc.id_pago where cccc.id_credito = c.cre_id and cccc.id_pago is not null and cccc.cuo_monto_capital > 0 and p.pag_fecha_pago <= EOMONTH(datefromparts(".$año.",".$mes.",1))) as decimal(19,2)) as [Saldo_Deudor_Calculado]
	FROM FBC_CREDITO C 
	INNER JOIN FBC_SOLICITUD S ON S.SOL_ID = C.ID_SOLICITUD 
	INNER JOIN FBC_PERSONA P ON P.PER_ID = S.ID_TITULAR 
	INNER JOIN FBC_LOCALIDAD L ON L.LOC_ID = P.ID_LOCALIDAD 
	inner join fbc_programa pr on pr.pro_id = s.id_programa_solicitado 
	WHERE 1 = 1 AND c.cre_fecha_efectivizacion  <= EOMONTH(datefromparts(".$año.",".$mes.",1))
	and S.id_programa_solicitado <> 43108
	and (C.CRE_MONTO_OTORGADO - isnull(((SELECT SUM(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS not NULL)),0)) > 0
	and isnull((SELECT count(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS NULL  and CU.cuo_monto_capital > 0 and CU.cuo_vencimiento_1 <=  EOMONTH(DATEFROMPARTS(".$año.",".$mes.",1))),0) = 1
	and s.sol_id not in (select id_solicitud from FBC_SEGUIMIENTO_SOLICITUD where sso_estado in ('CIERRE_ANTICIPADO','CIERRE_EXTRAORDINARIO','CERRADO'))
	) as X
	union all
	select '30-90 dias.' as Periodo, count(*) as Cantidad,  cast(sum([Calculo_Mora]) as decimal(19,2)) as Monto,
	cast(sum(Total_Pagado) as decimal(19,2)) as Pagado, cast(sum([Saldo_Deudor_Calculado]) as decimal(19,2)) as SALDO_DEUDOR
	from
	(
	SELECT 
	cast((C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * isnull((SELECT count(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS NULL and CU.cuo_monto_capital > 0 and CU.cuo_vencimiento_1 <=  EOMONTH(DATEFROMPARTS(".$año.",".$mes.",1))),0) as decimal(19,2)) as [Calculo_Mora],
	cast((C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * (select count(*) from fbc_cuota cccc  inner join fbc_pago p on p.pag_id = cccc.id_pago where cccc.id_credito = c.cre_id and cccc.id_pago is not null and p.pag_fecha_pago <= EOMONTH(datefromparts(".$año.",".$mes.",1))) as decimal(19,2)) as [Total_Pagado],
	cast(C.CRE_MONTO_OTORGADO - (C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * (select count(*) from fbc_cuota cccc  inner join fbc_pago p on p.pag_id = cccc.id_pago where cccc.id_credito = c.cre_id and cccc.id_pago is not null and cccc.cuo_monto_capital > 0 and p.pag_fecha_pago <= EOMONTH(datefromparts(".$año.",".$mes.",1))) as decimal(19,2)) as [Saldo_Deudor_Calculado]
	FROM FBC_CREDITO C 
	INNER JOIN FBC_SOLICITUD S ON S.SOL_ID = C.ID_SOLICITUD 
	INNER JOIN FBC_PERSONA P ON P.PER_ID = S.ID_TITULAR 
	INNER JOIN FBC_LOCALIDAD L ON L.LOC_ID = P.ID_LOCALIDAD 
	inner join fbc_programa pr on pr.pro_id = s.id_programa_solicitado 
	WHERE 1 = 1 AND c.cre_fecha_efectivizacion  <= EOMONTH(datefromparts(".$año.",".$mes.",1))
	and S.id_programa_solicitado <> 43108
	and (C.CRE_MONTO_OTORGADO - isnull(((SELECT SUM(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS not NULL)),0)) > 0
	and isnull((SELECT count(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS NULL  and CU.cuo_monto_capital > 0 and CU.cuo_vencimiento_1 <=  EOMONTH(DATEFROMPARTS(".$año.",".$mes.",1))),0) in (2,3)
	and s.sol_id not in (select id_solicitud from FBC_SEGUIMIENTO_SOLICITUD where sso_estado in ('CIERRE_ANTICIPADO','CIERRE_EXTRAORDINARIO','CERRADO'))
	) as X
	union all
	select '91 - 180 dias.' as Periodo, count(*) as Cantidad,  cast(sum([Calculo_Mora]) as decimal(19,2)) as Monto,
	cast(sum(Total_Pagado) as decimal(19,2)) as Pagado, cast(sum([Saldo_Deudor_Calculado]) as decimal(19,2)) as SALDO_DEUDOR
	from
	(
	SELECT 
	cast((C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * isnull((SELECT count(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS NULL and CU.cuo_monto_capital > 0 and CU.cuo_vencimiento_1 <=  EOMONTH(DATEFROMPARTS(".$año.",".$mes.",1))),0) as decimal(19,2)) as [Calculo_Mora],
	cast((C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * (select count(*) from fbc_cuota cccc  inner join fbc_pago p on p.pag_id = cccc.id_pago where cccc.id_credito = c.cre_id and cccc.id_pago is not null and p.pag_fecha_pago <= EOMONTH(datefromparts(".$año.",".$mes.",1))) as decimal(19,2)) as [Total_Pagado],
	cast(C.CRE_MONTO_OTORGADO - (C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * (select count(*) from fbc_cuota cccc  inner join fbc_pago p on p.pag_id = cccc.id_pago where cccc.id_credito = c.cre_id and cccc.id_pago is not null and cccc.cuo_monto_capital > 0 and p.pag_fecha_pago <= EOMONTH(datefromparts(".$año.",".$mes.",1))) as decimal(19,2)) as [Saldo_Deudor_Calculado]
	FROM FBC_CREDITO C 
	INNER JOIN FBC_SOLICITUD S ON S.SOL_ID = C.ID_SOLICITUD 
	INNER JOIN FBC_PERSONA P ON P.PER_ID = S.ID_TITULAR 
	INNER JOIN FBC_LOCALIDAD L ON L.LOC_ID = P.ID_LOCALIDAD 
	inner join fbc_programa pr on pr.pro_id = s.id_programa_solicitado 
	WHERE 1 = 1 AND c.cre_fecha_efectivizacion  <= EOMONTH(datefromparts(".$año.",".$mes.",1))
	and S.id_programa_solicitado <> 43108
	and (C.CRE_MONTO_OTORGADO - isnull(((SELECT SUM(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS not NULL)),0)) > 0
	and isnull((SELECT count(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS NULL  and CU.cuo_monto_capital > 0 and CU.cuo_vencimiento_1 <=  EOMONTH(DATEFROMPARTS(".$año.",".$mes.",1))),0) in (4,5,6)
	and s.sol_id not in (select id_solicitud from FBC_SEGUIMIENTO_SOLICITUD where sso_estado in ('CIERRE_ANTICIPADO','CIERRE_EXTRAORDINARIO','CERRADO'))
	) as X
	union all
	select '> 180 dias.' as Periodo, count(*) as Cantidad,  cast(sum([Calculo_Mora]) as decimal(19,2)) as Monto,
	cast(sum(Total_Pagado) as decimal(19,2)) as Pagado, cast(sum([Saldo_Deudor_Calculado]) as decimal(19,2)) as SALDO_DEUDOR
	from
	(
	SELECT 
	cast((C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * isnull((SELECT count(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS NULL and CU.cuo_monto_capital > 0 and CU.cuo_vencimiento_1 <=  EOMONTH(DATEFROMPARTS(".$año.",".$mes.",1))),0) as decimal(19,2)) as [Calculo_Mora],
	cast((C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * (select count(*) from fbc_cuota cccc  inner join fbc_pago p on p.pag_id = cccc.id_pago where cccc.id_credito = c.cre_id and cccc.id_pago is not null and p.pag_fecha_pago <= EOMONTH(datefromparts(".$año.",".$mes.",1))) as decimal(19,2)) as [Total_Pagado],
	cast(C.CRE_MONTO_OTORGADO - (C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * (select count(*) from fbc_cuota cccc  inner join fbc_pago p on p.pag_id = cccc.id_pago where cccc.id_credito = c.cre_id and cccc.id_pago is not null and cccc.cuo_monto_capital > 0 and p.pag_fecha_pago <= EOMONTH(datefromparts(".$año.",".$mes.",1))) as decimal(19,2)) as [Saldo_Deudor_Calculado]
	FROM FBC_CREDITO C 
	INNER JOIN FBC_SOLICITUD S ON S.SOL_ID = C.ID_SOLICITUD 
	INNER JOIN FBC_PERSONA P ON P.PER_ID = S.ID_TITULAR 
	INNER JOIN FBC_LOCALIDAD L ON L.LOC_ID = P.ID_LOCALIDAD 
	inner join fbc_programa pr on pr.pro_id = s.id_programa_solicitado 
	WHERE 1 = 1 AND c.cre_fecha_efectivizacion  <= EOMONTH(datefromparts(".$año.",".$mes.",1))
	and S.id_programa_solicitado <> 43108
	and (C.CRE_MONTO_OTORGADO - isnull(((SELECT SUM(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS not NULL)),0)) > 0
	and isnull((SELECT count(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS NULL  and CU.cuo_monto_capital > 0 and CU.cuo_vencimiento_1 <=  EOMONTH(DATEFROMPARTS(".$año.",".$mes.",1))),0) >= 7
	and s.sol_id not in (select id_solicitud from FBC_SEGUIMIENTO_SOLICITUD where sso_estado in ('CIERRE_ANTICIPADO','CIERRE_EXTRAORDINARIO','CERRADO'))
	) as X ";



echo '
         <div class="col-md-8 col-md-offset-3">
                    <div>
                        <div class="panel-body">
                             <form role="form" >
                                <fieldset>';
echo '                              <div class="col-md-7 col-md-offset-1">';
//echo $sql;

$stmt=sqlsrv_query($conn,$sql);
echo '<div class="text-center" >';
echo '    <h4>FONCAP</b></h4>';
echo '</div>';

echo '<table class="table" style="text-align: center;" border ="1">';
echo '  <tr style="text-align: center; background:#0099CC">';
echo '    <th style="text-align: center;">Periodo </th>';
echo '    <th style="text-align: center;">Cantidad </th>';
echo '    <th style="text-align: center;">Monto  </th>';
//echo '    <th style="text-align: center;">Pagado </th>';
//echo '    <th style="text-align: center;">Saldo Deudor </th>';
echo "  </tr>";
$suma = 0;
while ($row = sqlsrv_fetch_array($stmt))
{
  echo "<tr>";
  echo "<td>".$row['Periodo']."</td>";
  echo "<td>".$row['Cantidad']."</td>";
  echo '<td>$ '.$row['Monto'].'</td>';
//  echo "<td>".$row['Pagado']."</td>";
//  echo "<td>".$row['SALDO_DEUDOR']."</td>";
  $suma = $suma + ($row['SALDO_DEUDOR'] - $row['Monto']);
  echo "</tr>";
}
echo "<tr>";
echo "<td colspan=2>Cartera Cumplimiento Regular</td>";
echo "<td>$ ".$suma."</td>";
echo "</tr>";

echo "</table>";


echo '</div>';

echo '<div class="col-md-7 col-md-offset-1">';
echo "<br>";
echo '<div class="text-center" >';
echo '    <h4>NO FONCAP</b></h4>';
echo '</div>';

  sqlsrv_free_stmt($stmt);
  $stmt=sqlsrv_query($conn,$sqlNo);
echo '<table class="table" style="text-align: center;" border ="1">';
echo '  <tr style="text-align:center; background:#0099CC;">';
echo '    <th style="text-align: center;">Periodo </th>';
echo '    <th style="text-align: center;">Cantidad </th>';
echo '    <th style="text-align: center;">Monto  </th>';
echo "  </tr>";

$suma = 0;
while ($row = sqlsrv_fetch_array($stmt))
{
  echo "<tr>";
  echo "<td>".$row['Periodo']."</td>";
  echo "<td>".$row['Cantidad']."</td>";
  echo "<td>$ ".$row['Monto']."</td>";
  echo "</tr>";
  $suma = $suma + ($row['SALDO_DEUDOR'] - $row['Monto']);
}
echo "<tr>";
echo "<td colspan=2>Cartera Cumplimiento Regular</td>";
echo "<td>$ ".$suma."</td>";
echo "</tr>";

echo "</table>";
echo "<br>";
echo '</div>';

echo '       <div class="col-md-5 col-md-offset-2">
							<h6>(Generar el Excel puede demorar varios minutos)</b></h6><br>
			</div>';
echo '       <div class="col-md-5 col-md-offset-3">
               <button type="button" class="btn btn-light"><a href ="javascript:openPage('.$mes.','.$año.')">Descargar Excel <span class="glyphicon glyphicon-download-alt"></span></button>
	        </div>
          </fieldset> 
        </form>
      </div>
    </div> 
</div>


<script language="javascript" type="text/javascript">
openPage = function($mes,$año) {
location.href = "excel_anexoIII.php?mes="+$mes+"&año="+$año;
}
</script>'
;
?>