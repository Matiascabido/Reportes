<?php 

        session_start();

        ini_set('mssql.timeout',1000);
        set_time_limit(1000);

        $server = "fbcoprd.database.windows.net";
        $user = "adminfbco";
        $pwd="Fundacion#123";
        $dba="GestionCreditosFBCO";
        $concetinfo=array("Database" =>$dba , "UID" =>$user, "PWD"=>$pwd, "CharacterSet" => "UTF-8");
        $conn = sqlsrv_connect($server,$concetinfo);

        include 'PHPExcel-1.8/Classes/PHPExcel.php';
        include 'PHPExcel-1.8/Classes/PHPExcel/Writer/Excel2007.php';



        $mes = $_REQUEST['mes'];
        $anio = $_REQUEST['año'];

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
          /*Extraer datos de MYSQL*/

        $sql="SELECT C.ID_SOLICITUD [CREDITO], CC.CUO_NUMERO [CUOTA] ,CONCAT(S.SOL_PER_APELLIDO, ' ', S.SOL_PER_NOMBRE) AS [NOMBRE], CAST(CC.CUO_MONTO_CAPITAL as DECIMAL(19,4)) AS [CAPITAL], 
            SUM(CASE WHEN (DATEDIFF(D,CC.CUO_VENCIMIENTO_1, PA.PAG_FECHA_PAGO)) <= 0 
            THEN CUO_MONTO_INTERES_FINANCIERO_1 ELSE CASE WHEN (DATEDIFF(D,CC.CUO_VENCIMIENTO_2, PA.PAG_FECHA_PAGO)) <= 0  
            THEN CUO_MONTO_INTERES_FINANCIERO_2 ELSE CASE WHEN (DATEDIFF(D,CC.CUO_VENCIMIENTO_3, PA.PAG_FECHA_PAGO)) <= 0  
            THEN CUO_MONTO_INTERES_FINANCIERO_3 ELSE CUO_MONTO_INTERES_FINANCIERO_1+CUO_MONTO_INTERES_PUNITORIO END END END) AS INTERES
            ,PR.PRO_NOMBRE AS PROGRAMA, PA.PAG_FECHA_PAGO AS FECHA_PAGO 
            FROM FBC_CUOTA CC INNER JOIN FBC_CREDITO C ON C.CRE_ID = CC.ID_CREDITO 
            INNER JOIN FBC_PAGO PA ON PA.PAG_ID = CC.ID_PAGO INNER JOIN FBC_SOLICITUD S ON S.SOL_ID = C.ID_SOLICITUD 
            INNER JOIN FBC_PROGRAMA PR ON PR.PRO_ID = S.ID_PROGRAMA_SOLICITADO
            WHERE Month(PA.PAG_FECHA_PAGO) = $mes And Year(PA.PAG_FECHA_PAGO) = $anio 
            GROUP BY C.ID_SOLICITUD, CC.CUO_NUMERO, CONCAT(S.SOL_PER_APELLIDO, ' ', S.SOL_PER_NOMBRE), CC.CUO_MONTO_CAPITAL, PR.PRO_NOMBRE, PA.PAG_FECHA_PAGO";
        
        $stmt=sqlsrv_query($conn,$sql);   
//        $stmt->set_charset('utf8');
        $b = array("A","B","C","D","E", "F","G");
        $i = 0;
        $letra = '';
        $campos = array();
        $tam = array( "8","8","40","12","12","60","13");
        $tamf = "10";
        
        foreach( sqlsrv_field_metadata( $stmt ) as $fieldMetadata ) 
        {
          $letra = $b[$i];
          array_push( $campos,$fieldMetadata['Name']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue($letra.'1', $fieldMetadata['Name']);
          if($i < sizeof($tam))
            $objPHPExcel->getActiveSheet()->getColumnDimension($letra)->setWidth($tam[$i]);
          else
            $objPHPExcel->getActiveSheet()->getColumnDimension($letra)->setWidth($tamf);
          $i++; 
        }
        
        $cel = 2;
        /*Crédito Crédito Apellido y Nombres  Tipo y Numero de Documento  Capital Cuotas Pactadas Cant  Total Pagado  Programa

*/
        $tz_object = new DateTimeZone('Brazil/East');
        $datetime  = new DateTime();
        date_format($datetime, 'd/m/y');
        $datetime  ->setTimezone($tz_object);

        while ($row = sqlsrv_fetch_array($stmt))
        {
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('A'.$cel, $row['CREDITO']);   //credito
          $auxNombre  = $row['CUOTA'];
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('B'.$cel, $auxNombre);   //Apellido y Nombres
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('C'.$cel, $row['NOMBRE']);   //Apellido y Nombres
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('D'.$cel, $row['CAPITAL']);   //Sexo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('E'.$cel, $row['INTERES']);   //Interes
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('F'.$cel, $row['PROGRAMA']);   //Programa
          if (($row['FECHA_PAGO'] != null) || ($row['FECHA_PAGO'] != 0))
          {
            $datetime = $row['FECHA_PAGO'];
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('G'.$cel, PHPExcel_Shared_Date::PHPToExcel($datetime));
            $objPHPExcel->getActiveSheet()->getStyle('G'.$cel)->getNumberFormat()->setFormatCode('dd/mm/yyyy');
          }
          else
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('G'.$cel, '');   //Fecha de pago
//          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('G'.$cel, $row['FECHA_PAGO']);   //Domicilio
          $cel++;
        }
           /*Fin extracion de datos MYSQL*/

//        header('Content-Type: application/vnd.ms-excel');
        header("Content-type: application/vnd.ms-excel; charset=UTF-8" );
        header('Content-Disposition: attachment;filename="'.$anio.'-'.$mes.'-pagos_mes.xls"');
        header('Cache-Control: max-age=0');
         
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save('php://output');
 
?>
