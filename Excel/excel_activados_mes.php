<?php 

        include 'conn.php';
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

        $sql=" SELECT C.ID_SOLICITUD,  CONCAT(S.SOL_PER_APELLIDO,' ', S.SOL_PER_NOMBRE) AS TITULAR, PE.PER_NUM_DOC  AS DNI, PE.PER_CUIL_CUIT AS CUIT, 
             P.PRO_NOMBRE AS PROGRAMA, L.LOC_NOMBRE AS LOCALIDAD, D.DTO_DESCRIPCION AS DEPARTAMENTO, 
             (SELECT TOP 1 SS.SSO_ESTADO FROM FBC_SEGUIMIENTO_SOLICITUD SS WHERE SS.ID_SOLICITUD =S.SOL_ID ORDER BY SSO_ID DESC) AS ESTADO, 
             CAST(C.CRE_MONTO_OTORGADO AS DECIMAL(19,2)) AS CAPITAL,  CAST(C.CRE_MONTO_GASTOS_ADMINISTRATIVOS AS DECIMAL(19,2)) AS GASTOS, 
             CAST((SELECT SUM(CC.CUO_MONTO_INTERES_FINANCIERO_1) FROM FBC_CUOTA CC WHERE CC.ID_CREDITO = C.CRE_ID) AS DECIMAL(19,2)) AS INTERES, 
             C.CRE_CUOTAS_OTORGADAS AS [PLAN DE CUOTAS], 20 AS [TASA PUNITORIOS], 
             C.CRE_PERIODO_GRACIA_OTORGADO AS [GRACIA], CAST((SELECT TOP 1 CCC.CUO_MONTO_CAPITAL+CCC.CUO_MONTO_INTERES_FINANCIERO_1 FROM FBC_CUOTA CCC WHERE CCC.ID_CREDITO = C.CRE_ID AND CCC.CUO_MONTO_CAPITAL > 0) AS DECIMAL(19,2)) AS [VALOR CUOTA], 
             0 AS [CUOTAS EN MORA],  0 AS [CUOTAS ADEUDADAS], 0  AS MOROSIDAD, 0 AS  [DEUDA BASE]
             FROM FBC_CREDITO C INNER JOIN FBC_SOLICITUD S ON S.SOL_ID = C.ID_SOLICITUD 
             INNER JOIN FBC_PROGRAMA P ON P.PRO_ID = S.ID_PROGRAMA_SOLICITADO 
             INNER JOIN FBC_PERSONA PE ON PE.PER_ID = S.ID_TITULAR 
             INNER JOIN FBC_LOCALIDAD L ON L.LOC_ID = S.SOL_ID_LOCALIDAD 
             INNER JOIN FBC_DEPARTAMENTO D ON D.DTO_ID = L.ID_DEPARTAMENTO 
             INNER JOIN FBC_EMPRENDIMIENTO E ON E.EMP_ID = S.ID_EMPRENDIMIENTO 
             WHERE Year(C.CRE_FECHA_EFECTIVIZACION) = ".$anio." And Month(C.CRE_FECHA_EFECTIVIZACION) = ".$mes;
        
        $stmt=sqlsrv_query($conn,$sql);   
//        $stmt->set_charset('utf8');
        $b = array("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U");
        $i = 0;
        $letra = '';
        $campos = array();
        $tam = array( "13","40","9","12","64","24","18","20","8","8","9","16","17","7","13","17","19","12","12","12","50");
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
        /*ID_SOLICITUD  TITULAR DNI CUIT  PROGRAMA  LOCALIDAD DEPARTAMENTO  ESTADO  CAPITAL GASTOS  INTERES PLAN DE CUOTAS  TASA PUNITORIOS GRACIA  VALOR CUOTA CUOTAS EN MORA  
        CUOTAS ADEUDADAS  MOROSIDAD DEUDA BASE  ACTIVIDAD ACTIVIDAD MENOR


*/
        while ($row = sqlsrv_fetch_array($stmt))
        {
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('A'.$cel, $row['ID_SOLICITUD']);   //credito
          $auxNombre  = $row['TITULAR'];
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('B'.$cel, $auxNombre);   //Apellido y Nombres
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('C'.$cel, $row['DNI']);   //Apellido y Nombres
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('D'.$cel, $row['CUIT']);   //Sexo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('E'.$cel, $row['PROGRAMA']);   //Programa
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('F'.$cel,  $row['LOCALIDAD']);   //Domicilio
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('G'.$cel, $row['DEPARTAMENTO']);   //LOCALIDAD
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('H'.$cel, $row['ESTADO']);   //Provincia
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('I'.$cel, $row['CAPITAL']);   //Programa
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('J'.$cel, $row['GASTOS']);   //Programa
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('K'.$cel, $row['INTERES']);   //Programa
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('L'.$cel, $row['PLAN DE CUOTAS']);   //Programa
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('M'.$cel, $row['TASA PUNITORIOS']);   //Programa
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('N'.$cel, $row['GRACIA']);   //Programa
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('O'.$cel, $row['VALOR CUOTA']);   //Programa
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('P'.$cel, $row['CUOTAS EN MORA']);   //Programa
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('Q'.$cel, $row['CUOTAS ADEUDADAS']);   //Programa
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('R'.$cel, $row['MOROSIDAD']);   //Programa
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('S'.$cel, $row['DEUDA BASE']);   //Programa
          $cel++;
        }
           /*Fin extracion de datos MYSQL*/

//        header('Content-Type: application/vnd.ms-excel');
        header("Content-type: application/vnd.ms-excel; charset=UTF-8" );
        header('Content-Disposition: attachment;filename="'.$anio.'-'.$mes.'-cartera_incorporada.xls"');
        header('Cache-Control: max-age=0');
         
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save('php://output');


  
?>
