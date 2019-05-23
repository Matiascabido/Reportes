<?php 
        $mes = $_REQUEST['mes'];
        $anio = $_REQUEST['año'];

        include 'conn.php';

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
        $tz_object = new DateTimeZone('Brazil/East');
        $datetime  = new DateTime();
        $datetime2 = new DateTime();

        
        date_format($datetime, 'd/m/y');
        date_format($datetime2, 'd/m/y');


        $datetime  ->setTimezone($tz_object);
        $datetime2 ->setTimezone($tz_object);
       

        $sql="SELECT C.ID_SOLICITUD as [Nro Operacion], P.PER_APELLIDO + ' ' + P.PER_NOMBRE  as [Nombre y Apellido del Cliente], P.PER_FEC_NAC as [Fecha de Nacimiento],  P.PER_TIPO_DOC as [Tipo de Documento], P.PER_NUM_DOC as [Numero de Documento], concat(P.PER_CALLE,' ', P.PER_NUM_CALLE,' ', P.PER_PISO,' ',P.PER_DPTO, ' - ', 
            L.loc_nombre,' (' ,de.dto_nombre,')') as [Domicilio], P.PER_SEXO as [Genero], 
            C.CRE_FECHA_EFECTIVIZACION as [Fecha de Inicio de Microcredito],  
             C.CRE_MONTO_OTORGADO as [Suma Asegurada Inicial/Monto del Microcredito], (SELECT MAX(CU1.cuo_vencimiento_1) FROM FBC_CUOTA CU1 WHERE CU1.ID_CREDITO = C.CRE_ID) as [Fecha de Finalizacion del Microcredito],  
            isnull((C.CRE_MONTO_OTORGADO - isnull((SELECT SUM(CU.CUO_MONTO_CAPITAL)         FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS NOT NULL),0)),0) AS [Saldo Deudor de Capital al dia del reporte] ,  
            (select count(*) from FBC_CUOTA CU1 WHERE CU1.ID_CREDITO = C.CRE_ID AND CU1.ID_PAGO IS NULL AND CU1.cuo_vencimiento_1 < getdate()) as [CUOTAS ATRASADAS] 
            FROM FBC_CREDITO C INNER JOIN FBC_SOLICITUD S ON S.SOL_ID = C.ID_SOLICITUD INNER JOIN FBC_PERSONA P ON P.PER_ID = S.ID_TITULAR          
            LEFT OUTER JOIN FBC_LOCALIDAD L ON L.LOC_ID = S.SOL_ID_LOCALIDAD         
            LEFT OUTER JOIN FBC_DEPARTAMENTO DE ON DE.DTO_ID = L.ID_DEPARTAMENTO         
            WHERE C.CRE_FECHA_EFECTIVIZACION > DATEFROMPARTS (2016,6,1)         
            and isnull((C.CRE_MONTO_OTORGADO - isnull((SELECT SUM(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS NOT NULL),0)),0) > 0 
            AND C.CRE_SEGURO_INFORMADO IS NULL";
        
        $stmt=sqlsrv_query($conn,$sql);   

        $b = array("A","B","C","D","E","F","G","H","I","J","K","L");
        $i = 0;
        $letra = '';
        $campos = array();
        $tam = array( "9","40","27","9","15","6","50","12","12","12","12","12");
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


        while ($row = sqlsrv_fetch_array($stmt))
        {
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('A'.$cel, $row['Nro Operacion']);        
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('B'.$cel, $row['Nombre y Apellido del Cliente']);  
          
          
          if (($row['Fecha de Nacimiento'] != null) || ($row['Fecha de Nacimiento'] != 0))
          {
            $datetime2 = $row['Fecha de Nacimiento'];
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('C'.$cel, PHPExcel_Shared_Date::PHPToExcel($datetime2));
            $objPHPExcel->getActiveSheet()->getStyle('C'.$cel)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME);
          }
          else
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('C'.$cel, '');
          
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('D'.$cel, $row['Tipo de Documento']);      
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('E'.$cel, $row['Numero de Documento']);          
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('F'.$cel, $row['Domicilio']); 
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('G'.$cel, $row['Genero']);
          
          if (($row['Fecha de Inicio de Microcredito'] != null) || ($row['Fecha de Inicio de Microcredito'] != 0))
          {
            $datetime2 = $row['Fecha de Inicio de Microcredito'];
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('H'.$cel, PHPExcel_Shared_Date::PHPToExcel($datetime2));
            $objPHPExcel->getActiveSheet()->getStyle('H'.$cel)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME);
          }
          else
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('H'.$cel, '');
          
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('I'.$cel, $row['Suma Asegurada Inicial/Monto del Microcredito']);          
          
          if (($row['Fecha de Finalizacion del Microcredito'] != null) || ($row['Fecha de Finalizacion del Microcredito'] != 0))
          {
            $datetime2 = $row['Fecha de Finalizacion del Microcredito'];
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('J'.$cel, PHPExcel_Shared_Date::PHPToExcel($datetime2));
            $objPHPExcel->getActiveSheet()->getStyle('J'.$cel)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME);
          }
          else
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('J'.$cel, '');

          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('K'.$cel, $row['Saldo Deudor de Capital al dia del reporte']);  
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('L'.$cel, $row['CUOTAS ATRASADAS']);  
          $cel++;
        }

        header("Content-type: application/vnd.ms-excel; charset=UTF-8" );
        header('Content-Disposition: attachment;filename="'.$anio.'-'.$mes.'-REPORTE-SEGURO.xls"');
        header('Cache-Control: max-age=0');
         
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save('php://output');


  
?>
