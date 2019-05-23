<?php 

        include 'conn.php';
        include 'PHPExcel-1.8/Classes/PHPExcel.php';
        include 'PHPExcel-1.8/Classes/PHPExcel/Writer/Excel2007.php';

        ini_set('mssql.timeout',1000);
        set_time_limit(1000);

        $mes = $_REQUEST['mes'];
        $anio = $_REQUEST['año'];
        $dia = $_REQUEST['dia'];
        
        $mes2 = $_REQUEST['mes2'];
        $anio2 = $_REQUEST['año2'];
        $dia2 = $_REQUEST['dia2'];

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
        $datetime3 = new DateTime();
        $datetime4 = new DateTime();
        $datetime5 = new DateTime();
        $datetime6 = new DateTime();
        $datetime7 = new DateTime();
        $datetime8 = new DateTime();
        $datetime9 = new DateTime();
        $datetime10 = new DateTime();

        
        date_format($datetime, 'd/m/y');
        date_format($datetime2, 'd/m/y');
        date_format($datetime3, 'd/m/y');
        date_format($datetime4, 'd/m/y');
        date_format($datetime5, 'd/m/y');
        date_format($datetime6, 'd/m/y');
        date_format($datetime7, 'd/m/y');
        date_format($datetime8, 'd/m/y');
        date_format($datetime9, 'd/m/y');
        date_format($datetime10, 'd/m/y');


        $datetime  ->setTimezone($tz_object);
        $datetime2 ->setTimezone($tz_object);
        $datetime3 ->setTimezone($tz_object);
        $datetime4 ->setTimezone($tz_object);
        $datetime5 ->setTimezone($tz_object);
        $datetime6 ->setTimezone($tz_object);
        $datetime7 ->setTimezone($tz_object);
        $datetime8 ->setTimezone($tz_object);
        $datetime9 ->setTimezone($tz_object);
        $datetime10->setTimezone($tz_object);
         
       

        $sql="";
        
        $stmt=sqlsrv_query($conn,$sql);   

        $b = array("A","B","C","D","E","F","G","H","I");
        $i = 0;
        $letra = '';
        $campos = array();
        $tam = array( "9","40","27","9","15","6","50","12","12");
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
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('A'.$cel, $row['']);        
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('B'.$cel, $row['']);  
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('C'.$cel, $row['']);  
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('D'.$cel, $row['']);      
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('E'.$cel, $row['']);          
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('F'.$cel, $row['']); 
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('G'.$cel, $row['']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('H'.$cel, $row['']);         
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('I'.$cel, $row['']);          
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('J'.$cel, $row['']);        
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('K'.$cel, $row['']);  
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('L'.$cel, $row['']);  
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('M'.$cel, $row['']);      
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('N'.$cel, $row['']);          
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('O'.$cel, $row['']); 
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('P'.$cel, $row['']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('Q'.$cel, $row['']);         
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('R'.$cel, $row['']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('S'.$cel, $row['']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('T'.$cel, $row['']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('U'.$cel, $row['']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('V'.$cel, $row['']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('W'.$cel, $row['']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('X'.$cel, $row['']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('Y'.$cel, $row['']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('Z'.$cel, $row['']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AA'.$cel, $row['']);        
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AB'.$cel, $row['']);  
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AC'.$cel, $row['']);  
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AD'.$cel, $row['']);      
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AE'.$cel, $row['']);          
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AF'.$cel, $row['']); 
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AG'.$cel, $row['']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AH'.$cel, $row['']);         
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AI'.$cel, $row['']);          
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AJ'.$cel, $row['']);        
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AK'.$cel, $row['']);  
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AL'.$cel, $row['']);  
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AM'.$cel, $row['']);      
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AN'.$cel, $row['']);          
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AO'.$cel, $row['']); 
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AP'.$cel, $row['']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AQ'.$cel, $row['']);         
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AR'.$cel, $row['']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AS'.$cel, $row['']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AT'.$cel, $row['']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AU'.$cel, $row['']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AV'.$cel, $row['']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AW'.$cel, $row['']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AX'.$cel, $row['']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AY'.$cel, $row['']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AZ'.$cel, $row['']);
          if (($row[''] != null) || ($row[''] != 0))
          {
            $datetime= $row[''];
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue(''.$cel, PHPExcel_Shared_Date::PHPToExcel($datetime));
            $objPHPExcel->getActiveSheet()->getStyle(''.$cel)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME);
          }
          else
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue(''.$cel, '');
          if (($row[''] != null) || ($row[''] != 0))
          {
            $datetime2= $row[''];
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue(''.$cel, PHPExcel_Shared_Date::PHPToExcel($datetime2));
            $objPHPExcel->getActiveSheet()->getStyle(''.$cel)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME);
          }
          else
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue(''.$cel, '');

          if (($row[''] != null) || ($row[''] != 0))
          {
            $datetime3= $row[''];
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue(''.$cel, PHPExcel_Shared_Date::PHPToExcel($datetime3));
            $objPHPExcel->getActiveSheet()->getStyle(''.$cel)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME);
          }
          else
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue(''.$cel, '');

          
          $cel++;
        }

        header("Content-type: application/vnd.ms-excel; charset=UTF-8" );
        header('Content-Disposition: attachment;filename="'.$anio.'-'.$mes.'-Corte.xls"');
        header('Cache-Control: max-age=0');
         
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save('php://output');


  
?>
