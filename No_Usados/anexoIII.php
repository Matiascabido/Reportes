<?php 

        include 'conn.php';
         
        $server = "fbcoprd.database.windows.net";
        $user = "adminfbco";
        $pwd="Fundacion#123";
        $dba="GestionCreditosFBCO";
        $concetinfo=array("Database" =>$dba , "UID" =>$user, "PWD"=>$pwd, "CharacterSet" => "UTF-8");
        $conn = sqlsrv_connect($server,$concetinfo);

        include 'PHPExcel-1.8/Classes/PHPExcel.php';
        include 'PHPExcel-1.8/Classes/PHPExcel/Writer/Excel2007.php';

        ini_set('mssql.timeout',1000);
        set_time_limit(1000);

        $mes = $_REQUEST['mes'];
        $anio = $_REQUEST['año'];

        $objPHPExcel = new PHPExcel();
        $objPHPExcel->getProperties()->setCreator("Fundación Banco de Córdoba");
        $objPHPExcel->getProperties()->setLastModifiedBy("Fundación Banco de Córdoba");
        $objPHPExcel->getProperties()->setTitle("Reporte mensual");
        $objPHPExcel->getProperties()->setSubject("Asunto");
        $objPHPExcel->getProperties()->setDescription("Descripcion");
        $objPHPExcel->getActiveSheet()->setTitle('Reporte ');
        $objPHPExcel->setActiveSheetIndex(0);
        PHPExcel_Shared_Font::setAutoSizeMethod(PHPExcel_Shared_Font::AUTOSIZE_METHOD_EXACT); 
        $cacheMethod = PHPExcel_CachedObjectStorageFactory::cache_in_memory_gzip; 
        PHPExcel_Settings::setCacheStorageMethod($cacheMethod);
          /*Extraer datos de MYSQL*/

        $sql="SELECT C.ID_SOLICITUD as [Credito], 'SI' as [Credito Nuevo], P.PER_APELLIDO + ' ' + P.PER_NOMBRE  as [Apellido_y_Nombres], ";
        $sql.="P.PER_TIPO_DOC+' - '+P.PER_NUM_DOC as [Tipo y Numero de Documento], substring(P.per_sexo,1,1) as [Sexo],   ";
        $sql.="    concat(P.PER_CALLE,' ', P.PER_NUM_CALLE,' ', P.PER_PISO,' ',P.PER_DPTO) as [Domicilio],L.LOC_NOMBRE as [LOCALIDAD],    ";
        $sql.="    'Cordoba' as [Provincia],    ";
        $sql.="    (select top 1 aud_fecha_ins from fbc_seguimiento_solicitud sso where sso.id_solicitud = c.id_solicitud and sso.sso_estado = 'FIRMA_CONTRATO' order by sso_id desc) as [FECHA DEL PAGARE],";
        $sql.="    (select sum(cc.cuo_monto_capital+cc.cuo_monto_interes_financiero_1) from fbc_cuota cc where cc.id_credito = c.cre_id) as [Monto Pagare],    ";
        $sql.="    (select sum(cc.cuo_monto_capital+cc.cuo_monto_interes_financiero_1) from fbc_cuota cc where cc.id_credito = c.cre_id) as [Monto Credito],    ";
        $sql.="    (select top 1 ccc.cuo_monto_capital+ccc.cuo_monto_interes_financiero_1 from fbc_cuota ccc where ccc.id_Credito = c.cre_id and ccc.cuo_monto_capital > 0) as [Valor Cuota],   ";
        $sql.="    C.CRE_MONTO_OTORGADO as [Capital],  C.CRE_CUOTAS_OTORGADAS as [Cuotas Pactadas],    ";
        $sql.="    (select top 1 CU.cuo_vencimiento_1 from fbc_cuota cu where cu.id_credito = c.cre_id) as [FECHA 1er PAGO],";
        $sql.="    C.CRE_FECHA_FIN as [FECHA ULT. PAGO],   ";
        $sql.="    C.CRE_FECHA_FIN as [FECHA VTO.PAGARE],    ";
        $sql.="    'Mensual' as [PERIODICIDAD CUOTAS],    ";
        $sql.="    (select count(*) from fbc_cuota cccc  inner join fbc_pago p on p.pag_id = cccc.id_pago where cccc.id_credito = c.cre_id and cccc.id_pago is not null and p.pag_fecha_pago <= EOMONTH(datefromparts(".$anio.",".$mes.",1))) as [CUOTAS PAGADAS],   ";
        $sql.="    C.CRE_MONTO_OTORGADO - isnull(((SELECT SUM(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS not NULL)),0) AS [SALDO DEUDOR],    ";
        $sql.="    isnull((SELECT SUM(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS NULL and CU.cuo_monto_capital > 0 and CU.cuo_vencimiento_1 <=  EOMONTH(datefromparts(".$anio.",".$mes.",1))),0) AS [SALDO EN MORA],   ";
        $sql.="    isnull((SELECT count(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS NULL and CU.cuo_monto_capital > 0 and CU.cuo_vencimiento_1 <=  EOMONTH(datefromparts(".$anio.",".$mes.",1))),0) AS [CUOTAS EN MORA],  "; 
        $sql.="    cast((C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) as decimal(19,2)) as [CUOTA CAPITAL],    cast((C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * isnull((SELECT count(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS NULL and CU.cuo_monto_capital > 0 and CU.cuo_vencimiento_1 <=  EOMONTH(datefromparts(".$anio.",".$mes.",1))),0) as decimal(19,2)) as [Calculo Mora],    ";
        $sql.="    cast((C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * (select count(*) from fbc_cuota cccc  inner join fbc_pago p on p.pag_id = cccc.id_pago where cccc.id_credito = c.cre_id and cccc.id_pago is not null and p.pag_fecha_pago <= EOMONTH(datefromparts(".$anio.",".$mes.",1))) as decimal(19,2)) as [Total Pagado],    ";
        $sql.="    cast(C.CRE_MONTO_OTORGADO - (C.CRE_MONTO_OTORGADO /  C.CRE_CUOTAS_OTORGADAS) * (select count(*) from fbc_cuota cccc  inner join fbc_pago p on p.pag_id = cccc.id_pago where cccc.id_credito = c.cre_id and cccc.id_pago is not null and cccc.cuo_monto_capital > 0 and p.pag_fecha_pago <= EOMONTH(datefromparts(".$anio.",".$mes.",1))) as decimal(19,2)) as [Saldo Deudor Calculado],    ";
        $sql.="    pr.pro_nombre as [Programa], D.DTO_DESCRIPCION AS DEPARTAMENTO";
        $sql.="    FROM FBC_CREDITO C    INNER JOIN FBC_SOLICITUD S ON S.SOL_ID = C.ID_SOLICITUD    INNER JOIN FBC_PERSONA P ON P.PER_ID = S.ID_TITULAR    ";
        $sql.="    INNER JOIN FBC_LOCALIDAD L ON L.LOC_ID = P.ID_LOCALIDAD";    
        $sql.="    INNER JOIN FBC_DEPARTAMENTO D ON D.DTO_ID = L.ID_DEPARTAMENTO";
        $sql.="    inner join fbc_programa pr on pr.pro_id = s.id_programa_solicitado    ";
        $sql.="    WHERE 1 = 1 And c.cre_fecha_efectivizacion <= EOMONTH(datefromparts(".$anio.",".$mes.",1))    and (C.CRE_MONTO_OTORGADO - isnull(((SELECT SUM(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS not NULL)),0)) > 0";
        
        $stmt=sqlsrv_query($conn,$sql);   
//        $stmt->set_charset('utf8');
        $b = array("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA", "AB");
        $i = 0;
        $letra = '';
        $campos = array();
        $tipos  = array();
        $tam = array( "7","13","40","27","9","40","29","11","20","15","16","13","9","17","17","19","21","24","20","17","20","18","19","17","15","14","50","35");
        $tamf = "10";
        foreach( sqlsrv_field_metadata( $stmt ) as $fieldMetadata ) 
        {
          $letra = $b[$i];
          array_push( $campos,$fieldMetadata['Name']);
          array_push( $tipos,$fieldMetadata['Type']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue($letra.'1', $fieldMetadata['Name']);
          if($i < sizeof($tam))
            $objPHPExcel->getActiveSheet()->getColumnDimension($letra)->setWidth($tam[$i]);
          else
            $objPHPExcel->getActiveSheet()->getColumnDimension($letra)->setWidth($tamf);
          $i++; 
        }
        $cel = 2;
        /*Crédito 
        Credito Nuevo Apellido y Nombres  Tipo y Numero de Documento  Sexo  Domicilio LOCALIDAD Provincia FECHA DEL PAGARÉ  Monto Pagare  Monto Credito Valor Cuota Capital Cuotas Pactadas FECHA 1° PAGO FECHA ULT. PAGO FECHA VTO.PAGARE  PERIODICIDAD CUOTAS CUOTAS PAGADAS  SALDO DEUDOR  SALDO EN MORA CUOTAS EN MORA  CUOTA CAPITAL Calculo Mora  Total Pagado  Saldo Deudor Calculado  Programa
*/
        $tz_object = new DateTimeZone('Brazil/East');
        $datetime = new DateTime();
        $datetime2 = new DateTime();
        $datetime3 = new DateTime();
        $datetime4 = new DateTime();
        date_format($datetime, 'd/m/y');
        $datetime->setTimezone($tz_object);

        date_format($datetime2, 'd/m/y');
        $datetime2->setTimezone($tz_object);
        
        date_format($datetime3, 'd/m/y');
        $datetime3->setTimezone($tz_object);
        date_format($datetime4, 'd/m/y');
        $datetime4->setTimezone($tz_object);
        $auxNombre = '';

        while ($row = sqlsrv_fetch_array($stmt))
        {
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('A'.$cel, $row['Credito']);   //credito
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('B'.$cel, 'SI');   //Credito Nuevo
          $auxNombre  = $row['Apellido_y_Nombres'];
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('C'.$cel, $auxNombre);   //Apellido y Nombres
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('D'.$cel, $row['Tipo y Numero de Documento']);   //Apellido y Nombres
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('E'.$cel, $row['Sexo']);   //Sexo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('F'.$cel,  $row['Domicilio']);   //Domicilio
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('G'.$cel, $row['LOCALIDAD']);   //LOCALIDAD
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('H'.$cel, $row['Provincia']);   //Provincia
          if (($row['FECHA DEL PAGARE'] != null) || ($row['FECHA DEL PAGARE'] != 0))
          {
            $datetime= $row['FECHA DEL PAGARE'];
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('I'.$cel, PHPExcel_Shared_Date::PHPToExcel( $datetime));
            $objPHPExcel->getActiveSheet()->getStyle('I'.$cel)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME);
          }
          else
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('I'.$cel, '');

//          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('I'.$cel, );   //FECHA DEL PAGARÉ
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('J'.$cel, $row['Monto Pagare']);   //Monto Pagare
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('K'.$cel, $row['Monto Credito']);   //Monto Credito
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('L'.$cel, $row['Valor Cuota']);   //Valor Cuota
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('M'.$cel, $row['Capital']);   //Capital
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('N'.$cel, $row['Cuotas Pactadas']);   //Cuotas Pactadas
          
          $datetime2= $row['FECHA 1er PAGO'];
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('O'.$cel, PHPExcel_Shared_Date::PHPToExcel( $datetime2));
          $objPHPExcel->getActiveSheet()->getStyle('O'.$cel)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME);

//          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('O'.$cel, $row['FECHA 1er PAGO']);   //FECHA 1° PAGO
          $datetime3= $row['FECHA ULT. PAGO'];
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('P'.$cel, PHPExcel_Shared_Date::PHPToExcel( $datetime3));
          $objPHPExcel->getActiveSheet()->getStyle('P'.$cel)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME);

//          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('P'.$cel, $row['FECHA ULT. PAGO']);   //FECHA ULT. PAGO
          $datetime4= $row['FECHA VTO.PAGARE'];
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('Q'.$cel, PHPExcel_Shared_Date::PHPToExcel( $datetime4));
          $objPHPExcel->getActiveSheet()->getStyle('Q'.$cel)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME);
//          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('Q'.$cel, $row['FECHA VTO.PAGARE']);   //FECHA VTO.PAGARE
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('R'.$cel, 'Mensual');   //Cuotas Pactadas
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('S'.$cel, $row['CUOTAS PAGADAS']);   //Cuotas Pactadas
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('T'.$cel, $row['SALDO DEUDOR']);   //Cuotas Pactadas
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('U'.$cel, $row['SALDO EN MORA']);   //Cuotas Pactadas
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('V'.$cel, $row['CUOTAS EN MORA']);   //Cuotas Pactadas
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('W'.$cel, $row['CUOTA CAPITAL']);   //Programa
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('X'.$cel, $row['Calculo Mora']);   //Cuotas Pactadas
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('Y'.$cel, $row['Total Pagado']);   //Cuotas Pactadas
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('Z'.$cel, $row['Saldo Deudor Calculado']);   //Cuotas Pactadas
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AA'.$cel, $row['Programa']);   //Programa
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AB'.$cel, $row['Departamento']);   //Departamento
          $cel++;
        }
           /*Fin extracion de datos MYSQL*/

//        header('Content-Type: application/vnd.ms-excel');
        header("Content-type: application/vnd.ms-excel; charset=UTF-8" );
        header('Content-Disposition: attachment;filename="'.$anio.'-'.$mes.'-anexoIII.xls"');
        header('Cache-Control: max-age=0');
         
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save('php://output');

?>

