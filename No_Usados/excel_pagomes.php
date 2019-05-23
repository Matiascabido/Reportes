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
        $objPHPExcel->getActiveSheet()->setTitle('Reporte');
        $objPHPExcel->setActiveSheetIndex(0);
        PHPExcel_Shared_Font::setAutoSizeMethod(PHPExcel_Shared_Font::AUTOSIZE_METHOD_EXACT); 
        $cacheMethod = PHPExcel_CachedObjectStorageFactory::cache_in_memory_gzip; 
        PHPExcel_Settings::setCacheStorageMethod($cacheMethod);
          /*Extraer datos de MYSQL*/

        $sql ="SELECT C.ID_SOLICITUD [CREDITO], CC.CUO_NUMERO [CUOTA] ,CONCAT(S.SOL_PER_APELLIDO, ' ', S.SOL_PER_NOMBRE) AS [NOMBRE], 
                  CAST(CC.CUO_MONTO_CAPITAL as DECIMAL(19,4)) AS [CAPITAL], 
               (CASE WHEN (DATEDIFF(D,CC.CUO_VENCIMIENTO_1, PA.PAG_FECHA_PAGO)) <= 0 THEN CUO_MONTO_INTERES_FINANCIERO_1 
                    WHEN (DATEDIFF(D,CC.CUO_VENCIMIENTO_2, PA.PAG_FECHA_PAGO)) <= 0  THEN CUO_MONTO_INTERES_FINANCIERO_2  
                    WHEN (DATEDIFF(D,CC.CUO_VENCIMIENTO_3, PA.PAG_FECHA_PAGO)) <= 0  THEN CUO_MONTO_INTERES_FINANCIERO_3 ELSE
                     CUO_MONTO_INTERES_FINANCIERO_1+CUO_MONTO_INTERES_PUNITORIO END) AS INTERES --, SUM(CG.CGA_MONTO_GASTO) as GASTOS
              FROM FBC_CUOTA CC 
                  INNER JOIN FBC_CREDITO C ON C.CRE_ID = CC.ID_CREDITO 
                  INNER JOIN FBC_PAGO PA ON PA.PAG_ID = CC.ID_PAGO 
                  INNER JOIN FBC_SOLICITUD S ON S.SOL_ID = C.ID_SOLICITUD 
                  INNER JOIN FBC_CUOTA_GASTOS CG ON CG.ID_CUOTA = CC.CUO_ID
              WHERE Month(PA.PAG_FECHA_PAGO) = $mes And Year(PA.PAG_FECHA_PAGO) = $anio 
                  GROUP BY C.ID_SOLICITUD, CC.CUO_NUMERO, (CASE WHEN (DATEDIFF(D,CC.CUO_VENCIMIENTO_1, PA.PAG_FECHA_PAGO)) <= 0 THEN CUO_MONTO_INTERES_FINANCIERO_1 
                    WHEN (DATEDIFF(D,CC.CUO_VENCIMIENTO_2, PA.PAG_FECHA_PAGO)) <= 0  THEN CUO_MONTO_INTERES_FINANCIERO_2  
                    WHEN (DATEDIFF(D,CC.CUO_VENCIMIENTO_3, PA.PAG_FECHA_PAGO)) <= 0  THEN CUO_MONTO_INTERES_FINANCIERO_3 ELSE
                     CUO_MONTO_INTERES_FINANCIERO_1+CUO_MONTO_INTERES_PUNITORIO END) , CONCAT(S.SOL_PER_NOMBRE, ' ', S.SOL_PER_APELLIDO), CC.CUO_MONTO_CAPITAL";

/*
        $sql="SELECT C.ID_SOLICITUD [CREDITO], CC.CUO_NUMERO [CUOTA] ,CONCAT(S.SOL_PER_NOMBRE, ' ', S.SOL_PER_APELLIDO) AS [NOMBRE], 
          CAST(CC.CUO_MONTO_CAPITAL as DECIMAL(19,4)) AS [CAPITAL], SUM(CASE WHEN (DATEDIFF(D,CC.CUO_VENCIMIENTO_1, PA.PAG_FECHA_PAGO)) <= 0 THEN CUO_MONTO_INTERES_FINANCIERO_1 ELSE 
            CASE WHEN (DATEDIFF(D,CC.CUO_VENCIMIENTO_2, PA.PAG_FECHA_PAGO)) <= 0  THEN CUO_MONTO_INTERES_FINANCIERO_2 ELSE 
            CASE WHEN (DATEDIFF(D,CC.CUO_VENCIMIENTO_3, PA.PAG_FECHA_PAGO)) <= 0  THEN CUO_MONTO_INTERES_FINANCIERO_3 ELSE
             CUO_MONTO_INTERES_FINANCIERO_1+CUO_MONTO_INTERES_PUNITORIO END END END) AS INTERES 
          FROM FBC_CUOTA CC 
          INNER JOIN FBC_CREDITO C ON C.CRE_ID = CC.ID_CREDITO 
          INNER JOIN FBC_PAGO PA ON PA.PAG_ID = CC.ID_PAGO 
          INNER JOIN FBC_SOLICITUD S ON S.SOL_ID = C.ID_SOLICITUD 
          WHERE Month(PA.PAG_FECHA_PAGO) = $mes And Year(PA.PAG_FECHA_PAGO) = $anio 
          GROUP BY C.ID_SOLICITUD, CC.CUO_NUMERO, CONCAT(S.SOL_PER_NOMBRE, ' ', S.SOL_PER_APELLIDO), CC.CUO_MONTO_CAPITAL";
        */
        $stmt=sqlsrv_query($conn,$sql);   
//        $stmt->set_charset('utf8');
        //CREDITO CUOTA NOMBRE  CAPITAL INTERES

        $b = array("A","B","C","D","E");
        $i = 0;
        $letra = '';
        $campos = array();
        $tam = array( "11","9","80","12","12");
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
        while ($row = sqlsrv_fetch_array($stmt))
        {
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('A'.$cel, $row['CREDITO']);   //credito
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('B'.$cel, $row['CUOTA']);   //Apellido y Nombres
          $auxNombre  = $row['NOMBRE'];
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('C'.$cel, $auxNombre);   //Apellido y Nombres
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('D'.$cel, $row['CAPITAL']);   //Sexo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('E'.$cel,  $row['INTERES']);   //Domicilio
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
