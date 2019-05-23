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
         
        /*
        $sql="select s.sol_id as Solicitud, concat(s.sol_per_nombre,' ',s.sol_per_apellido) as Nombre, cast(s.sol_monto_solicitado as decimal(19,0)) as [Monto Solicitado], 
        l.loc_descripcion as Localidad, d.dto_descripcion as Departamento, p.pro_nombre as Programa, cast(c.cre_monto_otorgado as decimal(19,0)) as [Monto Otorgado],
        ss.sso_estado as Estado
        from FBC_SOLICITUD s
        inner join FBC_LOCALIDAD L on l.loc_id = s.sol_id_localidad
        inner join FBC_DEPARTAMENTO d on d.dto_id = l.id_departamento
        inner join FBC_PROGRAMA p on p.pro_id = s.id_programa_solicitado
        left outer join FBC_CREDITO c on c.id_solicitud = s.sol_id
        inner join FBC_SEGUIMIENTO_SOLICITUD ss on ss.id_solicitud = s.sol_id
        where ss.sso_estado in ('FIRMA_CONTRATO','EFECTIVIZACION','APROBADO','ANALISIS_CREDITICIO','SOLICITUD_INICIAL','EVALUACION_TECNICA','APROBACION_GERENCIA','ARMADO_LEGAJO')
        and ss.sso_id in (select max(sss.sso_id) from FBC_SEGUIMIENTO_SOLICITUD sss where sss.id_solicitud = s.sol_id)";
        */
         /*et.evt_monto_sugerido[Monto sugerido] en vez de s.sol_monto_solicitado */
        $sql = "select s.sol_id as Solicitud, concat(s.sol_per_apellido,' ',s.sol_per_nombre) as Nombre, cast(et.evt_monto_sugerido as decimal(19,0)) as [Monto Sugerido], 
        l.loc_descripcion as Localidad, d.dto_descripcion as Departamento, p.pro_nombre as Programa, cast(c.cre_monto_otorgado as decimal(19,0)) as [Monto Otorgado],
        ss.sso_estado as Estado
        from FBC_SOLICITUD s
        inner join FBC_LOCALIDAD L on l.loc_id = s.sol_id_localidad
        inner join FBC_DEPARTAMENTO d on d.dto_id = l.id_departamento
        inner join FBC_PROGRAMA p on p.pro_id = s.id_programa_solicitado
        left outer join FBC_CREDITO c on c.id_solicitud = s.sol_id
        inner join FBC_SEGUIMIENTO_SOLICITUD ss on ss.id_solicitud = s.sol_id
	    	inner join fbc_evaluacion_tecnica et on et.id_solicitud = s.sol_id
        where ss.sso_estado in ('FIRMA_CONTRATO','EFECTIVIZACION','APROBADO','ANALISIS_CREDITICIO','SOLICITUD_INICIAL','EVALUACION_TECNICA','APROBACION_GERENCIA','ARMADO_LEGAJO')
        and ss.sso_id in (select max(sss.sso_id) from FBC_SEGUIMIENTO_SOLICITUD sss where sss.id_solicitud = s.sol_id)";

        $stmt=sqlsrv_query($conn,$sql);   
        $b = array("A","B","C","D","E","F","G","H","I");
        $i = 0;
        $letra = '';
        $campos = array();
        $tam = array( "9","40","16","18","14","41","16","24");
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
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('A'.$cel, $row['Solicitud']);   //credito
          $auxNombre  = $row['Nombre'];
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('B'.$cel, $auxNombre);   //Apellido y Nombres
          if($row['Monto Sugerido'] == null){
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('C'.$cel, '0'); 
            }
            else
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('C'.$cel, $row['Monto Sugerido']);   //Apellido y Nombres 
         
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('D'.$cel, $row['Localidad']);   //Sexo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('E'.$cel, $row['Departamento']);   //Domicilio
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('F'.$cel, $row['Programa']);   //Domicilio
         
          if($row['Monto Otorgado'] == null){
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('G'.$cel, '0'); 
          }
          else
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('G'.$cel,$row['Monto Otorgado'] );   //Domicilio


          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('H'.$cel, $row['Estado']);   //Domicilio
          $cel++;
        }

        header("Content-type: application/vnd.ms-excel; charset=UTF-8" );
        header('Content-Disposition: attachment;filename="'.$anio.' -'.$mes.'- Solicitudes_PA.xls"');
        header('Cache-Control: max-age=0');
         
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save('php://output');
 
?>
