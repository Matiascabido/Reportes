<?php 

        include 'conn.php';
        include 'PHPExcel-1.8/Classes/PHPExcel.php';
        include 'PHPExcel-1.8/Classes/PHPExcel/Writer/Excel2007.php';

                
        $server = "fbcoprd.database.windows.net";
        $user = "adminfbco";
        $pwd="Fundacion#123";
        $dba="GestionCreditosFBCO";
        $concetinfo=array("Database" =>$dba , "UID" =>$user, "PWD"=>$pwd, "CharacterSet" => "UTF-8");
        $conn = sqlsrv_connect($server,$concetinfo);


        ini_set('mssql.timeout',1000);
        set_time_limit(1000);

        $mes = $_REQUEST['mes'];
        $anio = $_REQUEST['año'];

        $objPHPExcel = new PHPExcel();//Cuotas Otorgadas
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

    $sql = "SELECT C.ID_SOLICITUD, S.ID_TITULAR, E.EMP_NOMBRE AS [NOMBRE_EMPRENDIMIENTO], CONCAT(PE.PER_APELLIDO,' ',PE.PER_NOMBRE) AS [RAZON_TITULAR], 
        PE.PER_NUM_DOC AS [DNI_TITULAR], PE.PER_CUIL_CUIT AS [CUIT_TITULAR], L.LOC_NOMBRE AS LOCALIDAD, D.DTO_NOMBRE AS DEPARTAMENTO,  
        e.emp_actividad_principal as [Actividad],  do.dom_texto as [Subactividad], e.emp_barrio as [Barrio], 
        PR.PRO_ID AS CODIGO_PROGRAMA, PR.PRO_NOMBRE, cast(PR.PRO_MONTO_MAX as decimal(19,0)) AS MONTO_MAXIMO, PR.PRO_CUOTAS_MAX AS CUOTAS, 
        cast(PG.PGA_VALOR as decimal(19,2)) AS [INTERES ANUAL], C.CRE_ID AS ID_CREDITO, C.CRE_FECHA_EFECTIVIZACION,  
        C.AUD_USR_INS AS USUARIO_CREACION, cast(c.CRE_MONTO_OTORGADO as decimal(19,2)) as [Monto Credito],  C.CRE_CUOTAS_OTORGADAS AS [Cuotas Otorgadas],
         (SELECT TOP 1 cast(PG2.PGA_VALOR as decimal(19,2)) FROM FBC_PROGRAMA_GASTOS PG2 WHERE PG.ID_PROGRAMA = PR.PRO_ID AND PG2.PGA_GASTO = 'INTERES_PUNITORIO') AS [INTERES PUNITORIO], 
         C.CRE_PERIODO_GRACIA_OTORGADO AS PERIODO_GRACIA,  (SELECT TOP 1 SSO.SSO_ESTADO FROM FBC_SEGUIMIENTO_SOLICITUD SSO WHERE SSO.ID_SOLICITUD = S.SOL_ID ORDER BY SSO.SSO_ID DESC) AS [ESTADO], 
         C.CRE_FECHA_INICIO AS [FECHA INICIO], C.CRE_FECHA_FIN AS [FECHA FIN], 
         (SELECT MAX(P.PAG_FECHA_PAGO) FROM FBC_PAGO P WHERE P.PAG_ID IN (SELECT ID_PAGO FROM FBC_CUOTA CC WHERE CC.ID_CREDITO = C.CRE_ID AND ID_PAGO IS NOT NULL AND CC.CUO_MONTO_CAPITAL > 0 )) AS [ULTIMO_PAGO], 
         CAST(((SELECT COUNT(*) FROM FBC_CUOTA CC WHERE CC.ID_CREDITO = C.CRE_ID AND CC.ID_PAGO IS NULL AND CC.CUO_MONTO_CAPITAL > 0 AND CC.CUO_VENCIMIENTO_1 < GETDATE()) * (SELECT TOP 1 CCC.CUO_MONTO_CAPITAL FROM FBC_CUOTA CCC WHERE CCC.ID_CREDITO = C.CRE_ID AND CCC.ID_PAGO IS NULL AND CCC.CUO_MONTO_CAPITAL > 0  AND CCC.CUO_VENCIMIENTO_1 < GETDATE())) AS DECIMAL(19,2)) AS [MOROSIDAD CAPITAL],
          CAST(((SELECT COUNT(*) FROM FBC_CUOTA CC WHERE CC.ID_CREDITO = C.CRE_ID AND CC.ID_PAGO IS NULL AND CC.CUO_MONTO_CAPITAL > 0 AND CC.CUO_VENCIMIENTO_1 < GETDATE()) * (SELECT TOP 1 CCC.CUO_MONTO_CAPITAL+CCC.CUO_MONTO_INTERES_FINANCIERO_1 FROM FBC_CUOTA CCC WHERE CCC.ID_CREDITO = C.CRE_ID AND CCC.ID_PAGO IS NULL AND CCC.CUO_MONTO_CAPITAL > 0  AND CCC.CUO_VENCIMIENTO_1 < GETDATE())) AS DECIMAL(19,2)) AS [MOROSIDAD CAPITAL E INTERES] , 
         CAST((SELECT TOP 1 CCC.CUO_MONTO_CAPITAL+CCC.CUO_MONTO_INTERES_FINANCIERO_1 FROM FBC_CUOTA CCC WHERE CCC.ID_CREDITO = C.CRE_ID AND CCC.CUO_MONTO_CAPITAL > 0 ) AS DECIMAL(19,2))  AS [VALOR CUOTA] ,
         (SELECT COUNT(CCC.CUO_ID) FROM FBC_CUOTA CCC WHERE CCC.ID_CREDITO = C.CRE_ID AND CCC.CUO_VENCIMIENTO_1 < GETDATE() AND CCC.ID_PAGO IS NULL AND CCC.CUO_MONTO_CAPITAL > 0  AND CCC.CUO_MONTO_CAPITAL > 0 ) AS [CUOTASEN MORA] ,
         (SELECT COUNT(CCC.CUO_ID) FROM FBC_CUOTA CCC WHERE CCC.ID_CREDITO = C.CRE_ID AND CCC.ID_PAGO IS NULL AND CCC.CUO_MONTO_CAPITAL > 0 ) AS [CUOTAS ADEUDADAS] ,
          (SELECT MIN(CCC.CUO_VENCIMIENTO_1) FROM FBC_CUOTA CCC WHERE  CCC.ID_CREDITO = C.CRE_ID AND CCC.ID_PAGO IS NULL AND CCC.CUO_MONTO_CAPITAL > 0 AND CCC.CUO_VENCIMIENTO_1 < GETDATE() AND CCC.CUO_PLAN = 'NORMAL') AS [PRIMERA_CUOTA_VENCIDA ORIGINAL] , 
         (SELECT MIN(CCC.CUO_VENCIMIENTO_1) FROM FBC_CUOTA CCC WHERE  CCC.ID_CREDITO = C.CRE_ID AND CCC.ID_PAGO IS NULL AND CCC.CUO_MONTO_CAPITAL > 0 AND CCC.CUO_VENCIMIENTO_1 < GETDATE() AND CCC.CUO_PLAN = 'REFINANCIADO') AS [PRIMERA_CUOTA_VENCIDA REFINANCIACION] ,
          (SELECT MAX(CCC.CUO_VENCIMIENTO_1) FROM FBC_CUOTA CCC WHERE  CCC.ID_CREDITO = C.CRE_ID AND CCC.ID_PAGO IS NOT NULL AND CCC.CUO_PLAN = 'NORMAL') AS [ULTIMA_CUOTA_ABONADA ORIGINAL] ,
          (SELECT MAX(CCC.CUO_VENCIMIENTO_1) FROM FBC_CUOTA CCC WHERE  CCC.ID_CREDITO = C.CRE_ID AND CCC.ID_PAGO IS NOT NULL AND CCC.CUO_PLAN = 'REFINANCIADO') AS [ULTIMA_CUOTA_ABONADA REFINANCIACION] , 
         (SELECT MAX(P.PAG_FECHA_PAGO) FROM FBC_CUOTA CCC INNER JOIN FBC_PAGO P ON P.PAG_ID = CCC.ID_PAGO  WHERE  CCC.ID_CREDITO = C.CRE_ID AND CCC.ID_PAGO IS NOT NULL  AND CCC.CUO_PLAN = 'NORMAL') AS [FECHA_ULTIMO_PAGO ORIGINAL] ,
          (SELECT MAX(P.PAG_FECHA_PAGO) FROM FBC_CUOTA CCC INNER JOIN FBC_PAGO P ON P.PAG_ID = CCC.ID_PAGO  WHERE  CCC.ID_CREDITO = C.CRE_ID AND CCC.ID_PAGO IS NOT NULL  AND CCC.CUO_PLAN = 'REFINANCIADO') AS [FECHA_ULTIMO_PAGO REFINANCIACION] 
         ,(SELECT MAX(N.AUD_FECHA_INS) FROM FBC_SEGUIMIENTO_SOLICITUD N WHERE N.ID_SOLICITUD = S.SOL_ID) AS FECHA_ULTIMA_ACTUALIZACION ,
         (SELECT TOP 1 ES.SSO_OBSERVACIONES  FROM FBC_SEGUIMIENTO_SOLICITUD ES WHERE ES.ID_SOLICITUD = S.SOL_ID ORDER BY ES.SSO_ID DESC) AS ULTIMA_NOVEDAD ,
         CAST((SELECT SUM(CCC.CUO_MONTO_INTERES_FINANCIERO_1) FROM FBC_CUOTA CCC WHERE CCC.ID_CREDITO = C.CRE_ID AND CCC.CUO_PLAN =  'NORMAL') AS DECIMAL(19,2) )  AS [INTERES ORIGINAL] ,
         CAST((SELECT SUM(CCC.CUO_MONTO_INTERES_FINANCIERO_1) FROM FBC_CUOTA CCC WHERE CCC.ID_CREDITO = C.CRE_ID AND CCC.CUO_PLAN  = 'REFINANCIADO') AS DECIMAL(19,2) )  AS [INTERES REFINANCIADO] ,
         CAST((SELECT SUM(CCC.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CCC WHERE CCC.ID_CREDITO = C.CRE_ID AND CCC.ID_PAGO IS NULL AND CCC.CUO_MONTO_CAPITAL > 0 ) AS DECIMAL(19,2)) AS [DEUDA CAPITAL] , 
         S.SOL_PER_MAIL, SOL_PER_TEL_FIJO 
        FROM FBC_CREDITO C 
        INNER JOIN FBC_SOLICITUD S ON S.SOL_ID = C.ID_SOLICITUD 
        INNER JOIN FBC_EMPRENDIMIENTO E ON E.EMP_ID = S.ID_EMPRENDIMIENTO 
        INNER JOIN FBC_PERSONA PE ON PE.PER_ID = S.ID_TITULAR 
        INNER JOIN FBC_LOCALIDAD L ON L.LOC_ID = E.ID_LOCALIDAD 
        INNER JOIN FBC_DEPARTAMENTO D ON D.DTO_ID = L.ID_DEPARTAMENTO 
        INNER JOIN FBC_PROGRAMA PR ON PR.PRO_ID = S.ID_PROGRAMA_SOLICITADO 
        INNER JOIN FBC_PROGRAMA_GASTOS PG ON PG.ID_PROGRAMA = PR.PRO_ID AND PG.PGA_GASTO = 'INTERES_ANUAL' 
        LEFT OUTER JOIN FBC_DOMINIO DO ON DO.DOM_CLAVE = E.EMP_ACTIVIDAD_PRINCIPAL_DESC and DO.DOM_DOMINIO = 'ACTIVIDAD_SECUNDARIA_CLANAE' 
        ORDER BY S.SOL_ID";        
        $stmt=sqlsrv_query($conn,$sql);  
        $b = array("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS");
        $i = 0;
        $letra = '';
        $campos = array();
        $tipos  = array();
        $tam = array( "7","13","40","27","9","15","29","11","20","15","16","13","50","17","10","17","13","24","20","17","20","18","19","17","15","14","16" ,"15","16","13","9","17","17","19","21","24","20","17","20","18","19","17","15","14","50");
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

        $tz_object = new DateTimeZone('Brazil/East');
        $datetime = new DateTime();
        $datetime2 = new DateTime();
        $datetime3 = new DateTime();
        $datetime4 = new DateTime();
        $datetime5 = new DateTime();
        $datetime6 = new DateTime();
        $datetime7 = new DateTime();
        $datetime8 = new DateTime();
        $datetime9 = new DateTime();
        $datetime10 = new DateTime();
        $datetime11 = new DateTime();
        $datetime12 = new DateTime();

        date_format($datetime, 'd/m/y');
        $datetime->setTimezone($tz_object);

        date_format($datetime2, 'd/m/y');
        $datetime2->setTimezone($tz_object);
        
        date_format($datetime3, 'd/m/y');
        $datetime3->setTimezone($tz_object);
        date_format($datetime4, 'd/m/y');
        $datetime4->setTimezone($tz_object);

        date_format($datetime5, 'd/m/y');
        $datetime5->setTimezone($tz_object);

        date_format($datetime6, 'd/m/y');
        $datetime6->setTimezone($tz_object);

        date_format($datetime7, 'd/m/y');
        $datetime7->setTimezone($tz_object);

        date_format($datetime8, 'd/m/y');
        $datetime8->setTimezone($tz_object);

        date_format($datetime9, 'd/m/y');
        $datetime9->setTimezone($tz_object);

        date_format($datetime10, 'd/m/y');
        $datetime10->setTimezone($tz_object);

        date_format($datetime11, 'd/m/y');
        $datetime11->setTimezone($tz_object);

        date_format($datetime12, 'd/m/y');
        $datetime12->setTimezone($tz_object);

        while ($row = sqlsrv_fetch_array($stmt))
        {
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('A'.$cel, $row['ID_SOLICITUD']);   //credito
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('B'.$cel, $row['ID_TITULAR']);   //Credito Nuevo
          $auxNombre  = $row['NOMBRE_EMPRENDIMIENTO'];
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('C'.$cel, $auxNombre);   //Apellido y Nombres
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('D'.$cel, $row['RAZON_TITULAR']);   //Credito Nuevo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('E'.$cel, $row['DNI_TITULAR']);   //Credito Nuevo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('F'.$cel, $row['CUIT_TITULAR']);   //Credito Nuevo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('G'.$cel, $row['LOCALIDAD']);   //Credito Nuevo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('H'.$cel, $row['DEPARTAMENTO']);   //Credito Nuevo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('I'.$cel, $row['Actividad']);   //Credito Nuevo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('J'.$cel, $row['Subactividad']);   //Credito Nuevo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('K'.$cel, $row['Barrio']);   //Credito Nuevo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('L'.$cel, $row['CODIGO_PROGRAMA']);   //Credito Nuevo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('M'.$cel, $row['PRO_NOMBRE']);   //Credito Nuevo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('N'.$cel, $row['MONTO_MAXIMO']);   //Credito Nuevo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('O'.$cel, $row['CUOTAS']);   //Credito Nuevo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('P'.$cel, $row['INTERES ANUAL']);   //Credito Nuevo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('Q'.$cel, $row['ID_CREDITO']);   //Credito Nuevo

          if (($row['CRE_FECHA_EFECTIVIZACION'] != null) || ($row['CRE_FECHA_EFECTIVIZACION'] != 0))
          {
            $datetime= $row['CRE_FECHA_EFECTIVIZACION'];
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('R'.$cel, PHPExcel_Shared_Date::PHPToExcel($datetime));
            $objPHPExcel->getActiveSheet()->getStyle('R'.$cel)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME);
          }
          else
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('R'.$cel, '');

          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('S'.$cel, $row['USUARIO_CREACION']);   //Credito Nuevo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('T'.$cel, $row['Monto Credito']);   //Credito Nuevo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('U'.$cel, $row['Cuotas Otorgadas']);   //Credito Nuevo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('V'.$cel, $row['INTERES PUNITORIO']);   //Credito Nuevo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('W'.$cel, $row['PERIODO_GRACIA']);   //Credito Nuevo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('X'.$cel, $row['ESTADO']);   //Credito Nuevo

        if (($row['FECHA INICIO'] != null) || ($row['FECHA INICIO'] != 0))
          {
            $datetime2 = $row['FECHA INICIO'];
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('Y'.$cel, PHPExcel_Shared_Date::PHPToExcel($datetime2));
            $objPHPExcel->getActiveSheet()->getStyle('Y'.$cel)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME);
          }
          else
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('Y'.$cel, '');

        if (($row['FECHA FIN'] != null) || ($row['FECHA FIN'] != 0))
          {
            $datetime3 = $row['FECHA FIN'];
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('Z'.$cel, PHPExcel_Shared_Date::PHPToExcel($datetime3));
            $objPHPExcel->getActiveSheet()->getStyle('Z'.$cel)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME);
          }
          else
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('Z'.$cel, '');

        if (($row['ULTIMO_PAGO'] != null) || ($row['ULTIMO_PAGO'] != 0))
          {
            $datetime4 = $row['ULTIMO_PAGO'];
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AA'.$cel, PHPExcel_Shared_Date::PHPToExcel($datetime4));
            $objPHPExcel->getActiveSheet()->getStyle('AA'.$cel)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME);
          }
          else
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AA'.$cel, '');

          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AB'.$cel, $row['MOROSIDAD CAPITAL']);   //Credito Nuevo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AC'.$cel, $row['MOROSIDAD CAPITAL E INTERES']);   //Credito Nuevo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AD'.$cel, $row['VALOR CUOTA']);   //Credito Nuevo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AE'.$cel, $row['CUOTASEN MORA']);   //Credito Nuevo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AF'.$cel, $row['CUOTAS ADEUDADAS']);   //Credito Nuevo4

        if (($row['PRIMERA_CUOTA_VENCIDA ORIGINAL'] != null) || ($row['PRIMERA_CUOTA_VENCIDA ORIGINAL'] != 0))
          {
            $datetime6 = $row['PRIMERA_CUOTA_VENCIDA ORIGINAL'];
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AG'.$cel, PHPExcel_Shared_Date::PHPToExcel($datetime6));
            $objPHPExcel->getActiveSheet()->getStyle('AG'.$cel)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME);
          }
          else
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AG'.$cel, '');


        if (($row['PRIMERA_CUOTA_VENCIDA REFINANCIACION'] != null) || ($row['PRIMERA_CUOTA_VENCIDA REFINANCIACION'] != 0))
          {
            $datetime7 = $row['PRIMERA_CUOTA_VENCIDA REFINANCIACION'];
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AH'.$cel, PHPExcel_Shared_Date::PHPToExcel($datetime7));
            $objPHPExcel->getActiveSheet()->getStyle('AH'.$cel)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME);
          }
          else
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AH'.$cel, '');


        if (($row['ULTIMA_CUOTA_ABONADA ORIGINAL'] != null) || ($row['ULTIMA_CUOTA_ABONADA ORIGINAL'] != 0))
          {
            $datetime5 = $row['ULTIMA_CUOTA_ABONADA ORIGINAL'];
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AI'.$cel, PHPExcel_Shared_Date::PHPToExcel($datetime5));
            $objPHPExcel->getActiveSheet()->getStyle('AI'.$cel)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME);
          }
          else
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AI'.$cel, '');

        if (($row['ULTIMA_CUOTA_ABONADA REFINANCIACION'] != null) || ($row['ULTIMA_CUOTA_ABONADA REFINANCIACION'] != 0))
          {
            $datetime8 = $row['ULTIMA_CUOTA_ABONADA REFINANCIACION'];
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AJ'.$cel, PHPExcel_Shared_Date::PHPToExcel($datetime8));
            $objPHPExcel->getActiveSheet()->getStyle('AJ'.$cel)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME);
          }
          else
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AJ'.$cel, '');

        if (($row['FECHA_ULTIMO_PAGO ORIGINAL'] != null) || ($row['FECHA_ULTIMO_PAGO ORIGINAL'] != 0))
          {
            $datetime9 = $row['FECHA_ULTIMO_PAGO ORIGINAL'];
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AK'.$cel, PHPExcel_Shared_Date::PHPToExcel($datetime9));
            $objPHPExcel->getActiveSheet()->getStyle('AK'.$cel)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME);
          }
          else
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AK'.$cel, '');

        if (($row['FECHA_ULTIMO_PAGO REFINANCIACION'] != null) || ($row['FECHA_ULTIMO_PAGO REFINANCIACION'] != 0))
          {
            $datetime10 = $row['FECHA_ULTIMO_PAGO REFINANCIACION'];
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AL'.$cel, PHPExcel_Shared_Date::PHPToExcel($datetime10));
            $objPHPExcel->getActiveSheet()->getStyle('AL'.$cel)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME);
          }
          else
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AL'.$cel, '');

        if (($row['FECHA_ULTIMA_ACTUALIZACION'] != null) || ($row['FECHA_ULTIMA_ACTUALIZACION'] != 0))
          {
            $datetime11 = $row['FECHA_ULTIMA_ACTUALIZACION'];
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AM'.$cel, PHPExcel_Shared_Date::PHPToExcel($datetime11));
            $objPHPExcel->getActiveSheet()->getStyle('AM'.$cel)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME);
          }
          else
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AM'.$cel, '');


          
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AN'.$cel, $row['ULTIMA_NOVEDAD']);   //Credito Nuevo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AO'.$cel, $row['INTERES ORIGINAL']);   //Credito Nuevo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AP'.$cel, $row['INTERES REFINANCIADO']);   //Credito Nuevo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AQ'.$cel, $row['DEUDA CAPITAL']);   //Credito Nuevo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AR'.$cel, $row['SOL_PER_MAIL']);   //Credito Nuevo
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AS'.$cel, $row['SOL_PER_TEL_FIJO']);   //Credito Nuevo
          $cel++;
          

        }
           /*Fin extracion de datos MYSQL*/

        header("Content-type: application/vnd.ms-excel; charset=UTF-8" );
        header('Content-Disposition: attachment;filename="'.$anio.'-'.$mes.'-CarteraCompleta.xls"');
        header('Cache-Control: max-age=0');
         
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save('php://output');


?>



