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

        $sql="SELECT C.ID_SOLICITUD as [Credito],  P.PER_APELLIDO + ' ' + P.PER_NOMBRE  as [Apellido y Nombres],        
              P.PER_TIPO_DOC+' - '+P.PER_NUM_DOC as [Tipo y Numero de Documento],        C.CRE_MONTO_OTORGADO as [Capital],  
              C.CRE_CUOTAS_OTORGADAS as [Cuotas Pactadas],count(ccc.cuo_id) as Cant,        
              pr.pro_nombre as [Programa],
              cast(C.CRE_MONTO_OTORGADO / C.CRE_CUOTAS_OTORGADAS   as decimal(19,2)) as [VALOR CUOTA],
              cast((C.CRE_MONTO_OTORGADO / C.CRE_CUOTAS_OTORGADAS) * count(ccc.cuo_id)   as decimal(19,2)) as [TOTAL PAGADO]
              FROM FBC_CREDITO C    
              inner join fbc_cuota ccc on ccc.id_Credito = c.cre_id    
              INNER JOIN FBC_SOLICITUD S ON S.SOL_ID = C.ID_SOLICITUD    
              INNER JOIN FBC_PERSONA P ON P.PER_ID = S.ID_TITULAR   
              inner join fbc_programa pr on pr.pro_id = s.id_programa_solicitado    
              inner join fbc_pago pa on pa.pag_id = ccc.id_pago    
              where Year(pa.pag_fecha_pago) = $anio And Month(pa.pag_fecha_pago) = $mes   
              group by C.ID_SOLICITUD, P.PER_APELLIDO + ' ' + P.PER_NOMBRE, P.PER_TIPO_DOC+' - '+P.PER_NUM_DOC, C.CRE_MONTO_OTORGADO, 
               C.CRE_CUOTAS_OTORGADAS, pr.pro_nombre    
               order by c.id_solicitud";
        
        $stmt=sqlsrv_query($conn,$sql);   
//        $stmt->set_charset('utf8');
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
        /*Crédito Crédito Apellido y Nombres  Tipo y Numero de Documento  Capital Cuotas Pactadas Cant  Total Pagado  Programa

*/
        while ($row = sqlsrv_fetch_array($stmt))
        {
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('A'.$cel, $row['Credito']);   //credito
          $auxNombre  = $row['Apellido y Nombres'];
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('B'.$cel, $auxNombre);   //Apellido y Nombres
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('C'.$cel, $row['Tipo y Numero de Documento']);   //Apellido y Nombres
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('D'.$cel, $row['Capital']);   //Sexo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('E'.$cel,  $row['Cuotas Pactadas']);   //Domicilio
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('F'.$cel, $row['Cant']);   //LOCALIDAD
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('G'.$cel, $row['Programa']);   //Programa
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('H'.$cel, $row['VALOR CUOTA']);   //Programa
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('I'.$cel, $row['TOTAL PAGADO']);   //Programa
          $cel++;
        }
           /*Fin extracion de datos MYSQL*/

//        header('Content-Type: application/vnd.ms-excel');
        header("Content-type: application/vnd.ms-excel; charset=UTF-8" );
        header('Content-Disposition: attachment;filename="'.$anio.'-'.$mes.'-anexoI.xls"');
        header('Cache-Control: max-age=0');
         
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save('php://output');


  
?>
