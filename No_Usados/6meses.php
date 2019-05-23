<?php
  include_once('conn.php');
  
  include 'PHPExcel-1.8/Classes/PHPExcel.php';
  include 'PHPExcel-1.8/Classes/PHPExcel/Writer/Excel2007.php';

  ini_set('mssql.timeout', 1000);
  //        ini_set('mssql.charset', 'UTF-8');
  set_time_limit(1000);

  $objPHPExcel = new PHPExcel();
  $objPHPExcel->getProperties()->setCreator("Fundación Banco de Córdoba");
  $objPHPExcel->getProperties()->setLastModifiedBy("Fundación Banco de Córdoba");
  $objPHPExcel->getProperties()->setTitle("Reporte mensual");
  $objPHPExcel->getProperties()->setSubject("Asunto");
  $objPHPExcel->getProperties()->setDescription("Descripcion");
  $objPHPExcel->getActiveSheet()->setTitle('Reporte');
  $objPHPExcel->setActiveSheetIndex(0);
  PHPExcel_Shared_Font::setAutoSizeMethod(PHPExcel_Shared_Font::AUTOSIZE_METHOD_EXACT);
  $cacheMethod = PHPExcel_CachedObjectStorageFactory::cache_in_memory_gzip;
  PHPExcel_Settings::setCacheStorageMethod($cacheMethod);

  $sql = "
        Select 
        1 as Orden , year(P.pag_fecha_pago) as Anio, month(P.pag_fecha_pago) as Mes,  (select count(C.cuo_id) from fbc_cuota C where C.cuo_vencimiento_1 between DATEADD(d,1,EOMONTH(DATEADD(m,-13,getdate()))) and EOMONTH(DATEADD(m,-12,getdate()))) as [Del_Mes],
        count(P.pag_id) as Cantidad, cast(Sum(P.pag_monto_pagado) as decimal(19,2)) as [Monto Pagado], concat(year(P.pag_fecha_pago),'-',month(P.pag_fecha_pago)) as Periodo, 12 as Atras
        from FBC_PAGO P
        where P.pag_fecha_pago between DATEADD(d,1,EOMONTH(DATEADD(m,-13,getdate()))) and EOMONTH(DATEADD(m,-12,getdate()))
        group by  year(P.pag_fecha_pago), month(P.pag_fecha_pago), concat(year(P.pag_fecha_pago),'-',month(P.pag_fecha_pago))
        union 
        Select 
        2 as Orden , year(P.pag_fecha_pago) as Anio, month(P.pag_fecha_pago) as Mes, (select count(C.cuo_id) from fbc_cuota C where C.cuo_vencimiento_1 between DATEADD(d,1,EOMONTH(DATEADD(m,-12,getdate()))) and EOMONTH(DATEADD(m,-11,getdate()))) as [Del_Mes],
        count(P.pag_id) as Cantidad, cast(Sum(P.pag_monto_pagado) as decimal(19,2)) as [Monto Pagado], concat(year(P.pag_fecha_pago),'-',month(P.pag_fecha_pago)) as Periodo, 11 as Atras
        from FBC_PAGO P
        where P.pag_fecha_pago between DATEADD(d,1,EOMONTH(DATEADD(m,-12,getdate()))) and EOMONTH(DATEADD(m,-11,getdate()))
        group by  year(P.pag_fecha_pago), month(P.pag_fecha_pago), concat(year(P.pag_fecha_pago),'-',month(P.pag_fecha_pago))
        union 
        Select 
        3 as Orden , year(P.pag_fecha_pago) as Anio, month(P.pag_fecha_pago) as Mes, (select count(C.cuo_id) from fbc_cuota C where C.cuo_vencimiento_1 between DATEADD(d,1,EOMONTH(DATEADD(m,-11,getdate()))) and EOMONTH(DATEADD(m,-10,getdate()))) as [Del_Mes],
        count(P.pag_id) as Cantidad, cast(Sum(P.pag_monto_pagado) as decimal(19,2)) as [Monto Pagado], concat(year(P.pag_fecha_pago),'-',month(P.pag_fecha_pago)) as Periodo, 10 as Atras
        from FBC_PAGO P
        where P.pag_fecha_pago between DATEADD(d,1,EOMONTH(DATEADD(m,-11,getdate()))) and EOMONTH(DATEADD(m,-10,getdate()))
        group by year(P.pag_fecha_pago), month(P.pag_fecha_pago),  concat(year(P.pag_fecha_pago),'-',month(P.pag_fecha_pago))
        union 
        Select 
        4 as Orden , year(P.pag_fecha_pago) as Anio, month(P.pag_fecha_pago) as Mes, (select count(C.cuo_id) from fbc_cuota C where C.cuo_vencimiento_1 between DATEADD(d,1,EOMONTH(DATEADD(m,-10,getdate()))) and EOMONTH(DATEADD(m,-9,getdate()))) as [Del_Mes],
        count(P.pag_id) as Cantidad, cast(Sum(P.pag_monto_pagado) as decimal(19,2)) as [Monto Pagado], concat(year(P.pag_fecha_pago),'-',month(P.pag_fecha_pago)) as Periodo, 9 as Atras
        from FBC_PAGO P
        where P.pag_fecha_pago between DATEADD(d,1,EOMONTH(DATEADD(m,-10,getdate()))) and EOMONTH(DATEADD(m,-9,getdate()))
        group by  year(P.pag_fecha_pago), month(P.pag_fecha_pago), concat(year(P.pag_fecha_pago),'-',month(P.pag_fecha_pago))
        union 
        Select 
        5 as Orden , year(P.pag_fecha_pago) as Anio, month(P.pag_fecha_pago) as Mes, (select count(C.cuo_id) from fbc_cuota C where C.cuo_vencimiento_1 between DATEADD(d,1,EOMONTH(DATEADD(m,-9,getdate()))) and EOMONTH(DATEADD(m,-8,getdate()))) as [Del_Mes],
        count(P.pag_id) as Cantidad, cast(Sum(P.pag_monto_pagado) as decimal(19,2)) as [Monto Pagado], concat(year(P.pag_fecha_pago),'-',month(P.pag_fecha_pago)) as Periodo, 8 as Atras
        from FBC_PAGO P
        where P.pag_fecha_pago between DATEADD(d,1,EOMONTH(DATEADD(m,-9,getdate()))) and EOMONTH(DATEADD(m,-8,getdate()))
        group by  year(P.pag_fecha_pago), month(P.pag_fecha_pago), concat(year(P.pag_fecha_pago),'-',month(P.pag_fecha_pago))
        union 
        Select 
        6 as Orden , year(P.pag_fecha_pago) as Anio, month(P.pag_fecha_pago) as Mes, (select count(C.cuo_id) from fbc_cuota C where C.cuo_vencimiento_1 between DATEADD(d,1,EOMONTH(DATEADD(m,-8,getdate()))) and EOMONTH(DATEADD(m,-7,getdate()))) as [Del_Mes],
        count(P.pag_id) as Cantidad, cast(Sum(P.pag_monto_pagado) as decimal(19,2)) as [Monto Pagado], concat(year(P.pag_fecha_pago),'-',month(P.pag_fecha_pago)) as Periodo, 7 as Atras
        from FBC_PAGO P
        where P.pag_fecha_pago between DATEADD(d,1,EOMONTH(DATEADD(m,-8,getdate()))) and EOMONTH(DATEADD(m,-7,getdate()))
        group by  year(P.pag_fecha_pago), month(P.pag_fecha_pago), concat(year(P.pag_fecha_pago),'-',month(P.pag_fecha_pago))
        union 
        Select 
        7 as Orden , year(P.pag_fecha_pago) as Anio, month(P.pag_fecha_pago) as Mes, (select count(C.cuo_id) from fbc_cuota C where C.cuo_vencimiento_1 between DATEADD(d,1,EOMONTH(DATEADD(m,-7,getdate()))) and EOMONTH(DATEADD(m,-6,getdate()))) as [Del_Mes],
        count(P.pag_id) as Cantidad, cast(Sum(P.pag_monto_pagado) as decimal(19,2)) as [Monto Pagado], concat(year(P.pag_fecha_pago),'-',month(P.pag_fecha_pago)) as Periodo, 6 as Atras
        from FBC_PAGO P
        where P.pag_fecha_pago between DATEADD(d,1,EOMONTH(DATEADD(m,-7,getdate()))) and EOMONTH(DATEADD(m,-6,getdate()))
        group by  year(P.pag_fecha_pago), month(P.pag_fecha_pago), concat(year(P.pag_fecha_pago),'-',month(P.pag_fecha_pago))
        union 
        Select 
        8 as Orden , year(P.pag_fecha_pago) as Anio, month(P.pag_fecha_pago) as Mes, (select count(C.cuo_id) from fbc_cuota C where C.cuo_vencimiento_1 between DATEADD(d,1,EOMONTH(DATEADD(m,-6,getdate()))) and EOMONTH(DATEADD(m,-5,getdate()))) as [Del_Mes],
        count(P.pag_id) as Cantidad, cast(Sum(P.pag_monto_pagado) as decimal(19,2)) as [Monto Pagado], concat(year(P.pag_fecha_pago),'-',month(P.pag_fecha_pago)) as Periodo, 5 as Atras
        from FBC_PAGO P
        where P.pag_fecha_pago between DATEADD(d,1,EOMONTH(DATEADD(m,-6,getdate()))) and EOMONTH(DATEADD(m,-5,getdate()))
        group by year(P.pag_fecha_pago), month(P.pag_fecha_pago), concat(year(P.pag_fecha_pago),'-',month(P.pag_fecha_pago))
        union 
        Select 
        9 as Orden , year(P.pag_fecha_pago) as Anio, month(P.pag_fecha_pago) as Mes, (select count(C.cuo_id) from fbc_cuota C where C.cuo_vencimiento_1 between DATEADD(d,1,EOMONTH(DATEADD(m,-5,getdate()))) and EOMONTH(DATEADD(m,-4,getdate()))) as [Del_Mes],
        count(P.pag_id) as Cantidad, cast(Sum(P.pag_monto_pagado) as decimal(19,2)) as [Monto Pagado], concat(year(P.pag_fecha_pago),'-',month(P.pag_fecha_pago)) as Periodo, 4 as Atras
        from FBC_PAGO P
        where P.pag_fecha_pago between DATEADD(d,1,EOMONTH(DATEADD(m,-5,getdate()))) and EOMONTH(DATEADD(m,-4,getdate()))
        group by year(P.pag_fecha_pago), month(P.pag_fecha_pago), concat(year(P.pag_fecha_pago),'-',month(P.pag_fecha_pago))
        union 
        Select 
        10 as Orden , year(P.pag_fecha_pago) as Anio, month(P.pag_fecha_pago) as Mes, (select count(C.cuo_id) from fbc_cuota C where C.cuo_vencimiento_1 between DATEADD(d,1,EOMONTH(DATEADD(m,-4,getdate()))) and EOMONTH(DATEADD(m,-3,getdate()))) as [Del_Mes],
        count(P.pag_id) as Cantidad, cast(Sum(P.pag_monto_pagado) as decimal(19,2)) as [Monto Pagado], concat(year(P.pag_fecha_pago),'-',month(P.pag_fecha_pago)) as Periodo, 3 as Atras
        from FBC_PAGO P
        where P.pag_fecha_pago between DATEADD(d,1,EOMONTH(DATEADD(m,-4,getdate()))) and EOMONTH(DATEADD(m,-3,getdate()))
        group by year(P.pag_fecha_pago), month(P.pag_fecha_pago), concat(year(P.pag_fecha_pago),'-',month(P.pag_fecha_pago))
        union 
        Select 
        11 as Orden , year(P.pag_fecha_pago) as Anio, month(P.pag_fecha_pago) as Mes, (select count(C.cuo_id) from fbc_cuota C where C.cuo_vencimiento_1 between DATEADD(d,1,EOMONTH(DATEADD(m,-3,getdate()))) and EOMONTH(DATEADD(m,-2,getdate()))) as [Del_Mes],
        count(P.pag_id) as Cantidad, cast(Sum(P.pag_monto_pagado) as decimal(19,2)) as [Monto Pagado], concat(year(P.pag_fecha_pago),'-',month(P.pag_fecha_pago)) as Periodo, 2 as Atras
        from FBC_PAGO P
        where P.pag_fecha_pago between DATEADD(d,1,EOMONTH(DATEADD(m,-3,getdate()))) and EOMONTH(DATEADD(m,-2,getdate()))
        group by year(P.pag_fecha_pago), month(P.pag_fecha_pago), concat(year(P.pag_fecha_pago),'-',month(P.pag_fecha_pago))
        union 
        Select 
        12 as Orden , year(P.pag_fecha_pago) as Anio, month(P.pag_fecha_pago) as Mes, (select count(C.cuo_id) from fbc_cuota C where C.cuo_vencimiento_1 between DATEADD(d,1,EOMONTH(DATEADD(m,-2,getdate()))) and EOMONTH(DATEADD(m,-1,getdate()))) as [Del_Mes],
        count(P.pag_id) as Cantidad, cast(Sum(P.pag_monto_pagado) as decimal(19,2)) as [Monto Pagado], concat(year(P.pag_fecha_pago),'-',month(P.pag_fecha_pago)) as Periodo, 1 as Atras
        from FBC_PAGO P
        where P.pag_fecha_pago between DATEADD(d,1,EOMONTH(DATEADD(m,-2,getdate()))) and EOMONTH(DATEADD(m,-1,getdate()))
        group by year(P.pag_fecha_pago), month(P.pag_fecha_pago), concat(year(P.pag_fecha_pago),'-',month(P.pag_fecha_pago))
        order by Orden";

  $stmt = sqlsrv_query($conn, $sql);
  $lista = array();
  while ($row = sqlsrv_fetch_array($stmt)) {

    array_push($row, 0);
    array_push($row, 0);
    array_push($row, 0);
    array_push($row, 0);
    array_push($row, '');
    array_push($lista, $row);
  }
  sqlsrv_free_stmt($stmt);

  for ($i=0; $i< sizeof($lista); $i++){
    $sql = " Select year(cc.cuo_vencimiento_1) as Anio, month(cc.cuo_vencimiento_1) as Mes, cast(sum(cc.cuo_monto_capital+cc.cuo_monto_interes_financiero_1) as decimal(19,2)) as [Monto Cuotas]
            from  FBC_CUOTA cc 
            where concat(year(cc.cuo_vencimiento_1),'-',month(cc.cuo_vencimiento_1)) = '".$lista[$i]['Periodo']."'";
    $sql .= " group by year(cc.cuo_vencimiento_1), month(cc.cuo_vencimiento_1)
            union
            Select year(cc.cuo_vencimiento_1) as Anio, month(cc.cuo_vencimiento_1) as Mes,cast(sum(cc.cuo_monto_capital+cc.cuo_monto_interes_financiero_1) as decimal(19,2)) as [Monto Cuotas]
            from  FBC_CUOTA cc 
            where concat(year(cc.cuo_vencimiento_1),'-',month(cc.cuo_vencimiento_1)) = '".($lista[$i]['Anio']+1).'-'.$lista[$i]['Mes']."'";
      $sql .= " group by year(cc.cuo_vencimiento_1), month(cc.cuo_vencimiento_1)
             order by anio ";
    $stmt = sqlsrv_query($conn, $sql);
    if ($row = sqlsrv_fetch_array($stmt)) {

      if ($row['Monto Cuotas'] >= $lista[$i]['Monto Pagado'] )
      {
        $lista[$i][0] = $row['Monto Cuotas']; //CUOTAS
        $lista[$i][1] = $lista[$i]['Monto Pagado']; //PAGADO
        $lista[$i][2] = ($lista[$i]['Monto Pagado']*100/$row['Monto Cuotas']); 
        $row = sqlsrv_fetch_array($stmt);
        $lista[$i][3] = $row['Monto Cuotas'];
      }
      else
      {
        $lista[$i][0] = $row['Monto Cuotas'];
        $lista[$i][1] = $row['Monto Cuotas'];
        $lista[$i][2] = 100; 
        $row = sqlsrv_fetch_array($stmt);
        $lista[$i][3] = $row['Monto Cuotas'];
      }
      $lista[$i][4] = $row['Anio'].'-'.$row['Mes']; 

      sqlsrv_free_stmt($stmt);
      /*
      echo $lista[$i]['Periodo'].'-  Cuanto $ '.$lista[$i][0].' -> Pagado $ '.$lista[$i]['Monto Pagado'].' - Porc:'.number_format(($lista[$i]['Monto Pagado']*100/$lista[$i][0]),2).' - Prox: '.number_format($lista[$i][1],2); 
      echo ' - el: '.number_format($lista[$i][2],2);
      echo ' - el NRO: '.number_format($lista[$i][3],2);
      echo '<br>';
      */
    }
  }
  $b = array("A", "B", "C", "D", "E", "F", "G","H");
  $i = 0;
  $letra = '';
  $campos = array("Periodo","A Cobrar","Cobrado","Porc. %","Periodo Fut.","A Cobrar","Porc. %"," Esperado");
  $tam = array("8", "12", "12", "8", "8", "12", "8", "12");
  $tamf = "10";

  for($i=0; $i<8 ;$i++){
    $letra = $b[$i];
      $objPHPExcel->setActiveSheetIndex(0)->setCellValue($letra . '1', $campos[$i]);
      if ($i < sizeof($tam)) {
          $objPHPExcel->getActiveSheet()->getColumnDimension($letra)->setWidth($tam[$i]);
      } else {
          $objPHPExcel->getActiveSheet()->getColumnDimension($letra)->setWidth($tamf);
      }
  }
  $cel = 2;
  //for($i=1; $i<=sizeof($lista))
  $i =0;
  foreach ($lista as $elemento) {
    # code...
    //number_format(number,decimals,decimalpoint,separator)
    $objPHPExcel->setActiveSheetIndex(0)->setCellValue('A' . $cel, $lista[$i]['Periodo']); 
    $objPHPExcel->setActiveSheetIndex(0)->setCellValue('B' . $cel, number_format($lista[$i][0],2,'.','')); 
    $objPHPExcel->setActiveSheetIndex(0)->setCellValue('C' . $cel, number_format($lista[$i][1],2,'.',''));
    $objPHPExcel->setActiveSheetIndex(0)->setCellValue('D' . $cel, number_format($lista[$i][2],2,'.',''));
    $objPHPExcel->setActiveSheetIndex(0)->setCellValue('E' . $cel, $lista[$i][4],2);
    $objPHPExcel->setActiveSheetIndex(0)->setCellValue('F' . $cel, number_format($lista[$i][3],2,'.',''));
    $objPHPExcel->setActiveSheetIndex(0)->setCellValue('G' . $cel, number_format($lista[$i][2],2,'.',''));
    $objPHPExcel->setActiveSheetIndex(0)->setCellValue('H' . $cel, '=F'.$cel.'*G'.$cel.'/100');
    $cel++;
    $i++;    
  }
  header("Content-type: application/vnd.ms-excel; charset=UTF-8");
  header('Content-Disposition: attachment;filename="delangel.xls"');
  header('Cache-Control: max-age=0');

  $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
  $objWriter->save('php://output');
  $cel = 2;


?>