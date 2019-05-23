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
         
       

        $sql="SELECT C.ID_SOLICITUD,  'TITULAR' AS CARACTER_TITULAR, CONCAT(PE.PER_APELLIDO,' ',PE.PER_NOMBRE) AS [RAZON_TITULAR], PE.PER_NUM_DOC AS [DNI_TITULAR], PE.PER_CUIL_CUIT AS [CUIT_TITULAR], CONCAT(PE.PER_CALLE,' ',PE.PER_NUM_CALLE,' - ', PE.PER_COD_POSTAL, ' - ', L.LOC_NOMBRE,' (', D.DTO_NOMBRE,')') as DOMICILIO_TITULAR,  
              PR.PRO_NOMBRE, CAST(PR.PRO_MONTO_MAX AS DECIMAL(19,0)) AS MONTO_MAXIMO, PR.PRO_CUOTAS_MAX AS CUOTAS, 
                  CAST(C.CRE_MONTO_OTORGADO AS DECIMAL(19,2)) AS [MONTO CREDITO],  C.CRE_CUOTAS_OTORGADAS AS [CUOTAS OTORGADAS],
                   C.CRE_PERIODO_GRACIA_OTORGADO AS PERIODO_GRACIA,  (SELECT TOP 1 SSO.SSO_ESTADO FROM FBC_SEGUIMIENTO_SOLICITUD SSO WHERE SSO.ID_SOLICITUD = S.SOL_ID ORDER BY SSO.SSO_ID DESC) AS [ESTADO], 
                   (SELECT MAX(P.PAG_FECHA_PAGO) FROM FBC_PAGO P WHERE P.PAG_ID IN (SELECT ID_PAGO FROM FBC_CUOTA CC WHERE CC.ID_CREDITO = C.CRE_ID AND ID_PAGO IS NOT NULL AND CC.CUO_MONTO_CAPITAL > 0 )) AS [ULTIMO_PAGO], 
                   CAST(((SELECT COUNT(*) FROM FBC_CUOTA CC WHERE CC.ID_CREDITO = C.CRE_ID AND CC.ID_PAGO IS NULL AND CC.CUO_MONTO_CAPITAL > 0 AND CC.CUO_VENCIMIENTO_1 < GETDATE()) * (SELECT TOP 1 CCC.CUO_MONTO_CAPITAL FROM FBC_CUOTA CCC WHERE CCC.ID_CREDITO = C.CRE_ID AND CCC.ID_PAGO IS NULL AND CCC.CUO_MONTO_CAPITAL > 0 AND CCC.CUO_VENCIMIENTO_1 < GETDATE())) AS DECIMAL(19,2)) AS [MOROSIDAD CAPITAL],
               (SELECT COUNT(CCC.CUO_ID) FROM FBC_CUOTA CCC WHERE CCC.ID_CREDITO = C.CRE_ID AND CCC.ID_PAGO IS NULL AND CCC.CUO_MONTO_CAPITAL > 0 ) AS [CUOTAS ADEUDADAS] ,
(SELECT COUNT(CCC.CUO_ID) FROM FBC_CUOTA CCC WHERE CCC.ID_CREDITO = C.CRE_ID AND CCC.ID_PAGO IS NULL AND CCC.CUO_MONTO_CAPITAL > 0 AND CCC.CUO_VENCIMIENTO_1 < GETDATE()) AS [CUOTAS VENCIDAS] ,                   (SELECT MAX(N.AUD_FECHA_INS) FROM FBC_SEGUIMIENTO_SOLICITUD N WHERE N.ID_SOLICITUD = S.SOL_ID) AS FECHA_ULTIMA_ACTUALIZACION ,
                   (SELECT TOP 1 ES.SSO_OBSERVACIONES  FROM FBC_SEGUIMIENTO_SOLICITUD ES WHERE ES.ID_SOLICITUD = S.SOL_ID ORDER BY ES.SSO_ID DESC) AS ULTIMA_NOVEDAD ,
                   CAST((SELECT SUM(CCC.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CCC WHERE CCC.ID_CREDITO = C.CRE_ID AND CCC.ID_PAGO IS NULL AND CCC.CUO_MONTO_CAPITAL > 0 ) AS DECIMAL(19,2)) AS [DEUDA CAPITAL] , 
                   S.SOL_PER_MAIL, SOL_PER_TEL_FIJO, SOL_PER_TEL_CEL,
              'GARANTE' AS CARACTER_GARANTE, 
              ISNULL((SELECT TOP 1  CONCAT(PER2.PER_NOMBRE,' ',PER2.PER_APELLIDO) FROM FBC_SOLICITUD S2
              INNER JOIN FBC_SOLICITUD_GARANTIA SG2 ON SG2.ID_SOLICITUD = S2.SOL_ID 
              INNER JOIN FBC_GARANTIA G2 ON G2.GAR_ID = SG2.ID_GARANTE
              INNER JOIN FBC_PERSONA PER2 ON PER2.PER_ID = G2.ID_PERSONA
              WHERE S2.SOL_ID = S.SOL_ID),'')  AS NOMBRE_GARANTE, 
              ISNULL((SELECT TOP 1  CONCAT(PER2.PER_CALLE, ' ',PER2.PER_NUM_CALLE,' ',PER2.PER_PISO,' ', PER2.PER_DPTO, ' - ' ,REPLACE(PER2.PER_COD_POSTAL,'SIN_DATOS',' '), ' - ', L2.LOC_NOMBRE, '(',D2.DTO_NOMBRE,')')  FROM FBC_SOLICITUD S2
              INNER JOIN FBC_SOLICITUD_GARANTIA SG2 ON SG2.ID_SOLICITUD = S2.SOL_ID 
              INNER JOIN FBC_GARANTIA G2 ON G2.GAR_ID = SG2.ID_GARANTE
              INNER JOIN FBC_PERSONA PER2 ON PER2.PER_ID = G2.ID_PERSONA
              INNER JOIN FBC_LOCALIDAD L2 ON L2.LOC_ID = PER2.ID_LOCALIDAD
              INNER JOIN FBC_DEPARTAMENTO D2 ON D2.DTO_ID = L2.ID_DEPARTAMENTO WHERE S2.SOL_ID = S.SOL_ID),'')  AS DOMICILIO_GARANTE,
              (SELECT COUNT(*) FROM  FBC_SOLICITUD_GARANTIA SG2 
              INNER JOIN FBC_GARANTIA G2 ON G2.GAR_ID = SG2.ID_GARANTE
              WHERE SG2.ID_SOLICITUD = S.SOL_ID)  AS GARANTES, 
              REPLACE(ISNULL((SELECT TOP 1  CONCAT(PER2.PER_TEL_FIJO, ' / ',PER2.PER_TEL_CEL)  FROM FBC_SOLICITUD S2
              INNER JOIN FBC_SOLICITUD_GARANTIA SG2 ON SG2.ID_SOLICITUD = S2.SOL_ID 
              INNER JOIN FBC_GARANTIA G2 ON G2.GAR_ID = SG2.ID_GARANTE
              INNER JOIN FBC_PERSONA PER2 ON PER2.PER_ID = G2.ID_PERSONA
              WHERE S2.SOL_ID = S.SOL_ID), ''),'0 / 0','')  AS TELEFONOS
              , CONCAT(ATM.ATM_NOMBRE, ' ', ATM.ATM_APELLIDO) AS ATM
              , C.CRE_FECHA_EFECTIVIZACION as FECHA_EFECTIVIZACION, L.LOC_ID, L.LOC_NOMBRE
                  FROM FBC_CREDITO C 
                  INNER JOIN FBC_SOLICITUD S ON S.SOL_ID = C.ID_SOLICITUD 
                  INNER JOIN FBC_EMPRENDIMIENTO E ON E.EMP_ID = S.ID_EMPRENDIMIENTO 
                  INNER JOIN FBC_PERSONA PE ON PE.PER_ID = S.ID_TITULAR 
                  INNER JOIN FBC_LOCALIDAD L ON L.LOC_ID = PE.ID_LOCALIDAD 
                  INNER JOIN FBC_DEPARTAMENTO D ON D.DTO_ID = L.ID_DEPARTAMENTO 
                  INNER JOIN FBC_PROGRAMA PR ON PR.PRO_ID = S.ID_PROGRAMA_SOLICITADO 
              LEFT OUTER JOIN FBC_ATM ATM  ON ATM.ATM_ID = S.ID_ATM";
        
        $stmt=sqlsrv_query($conn,$sql);   
        /*
        ID_SOLICITUD  CARACTER_TITULAR  RAZON_TITULAR DNI_TITULAR CUIT_TITULAR  LOCALIDAD DEPARTAMENTO  PRO_NOMBRE  MONTO_MAXIMO  CUOTAS  MONTO CREDITO CUOTAS OTORGADAS  
        PERIODO_GRACIA  ESTADO  ULTIMO_PAGO MOROSIDAD CAPITAL CUOTAS ADEUDADAS  FECHA_ULTIMA_ACTUALIZACION  ULTIMA_NOVEDAD  DEUDA CAPITAL 
        SOL_PER_MAIL  SOL_PER_TEL_FIJO  CARACTER_GARANTE  NOMBRE_GARANTE  DOMICILIO_GARANTE GARANTES  TELEFONOS ATM
      */

        $b = array("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD","AE","AF");
        $i = 0;
        $letra = '';
        $campos = array();
        $tam = array("13", "18", "40", "13", "13", "57", "50", "17", "9",  "16", "20", "17", "30", "14", "20", "19", "30", "60", "15", "40", "18","18", "19", "39", "65", "10", "30", "51","20","10","30");
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
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('A'.$cel, $row['ID_SOLICITUD']);        
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('B'.$cel, $row['CARACTER_TITULAR']);  
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('C'.$cel, $row['RAZON_TITULAR']);  
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('D'.$cel, $row['DNI_TITULAR']);      
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('E'.$cel, $row['CUIT_TITULAR']);          
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('F'.$cel, $row['DOMICILIO_TITULAR']); 
//          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('F'.$cel, $row['CALLE']); 
//          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('G'.$cel, $row['LOCALIDAD']); 
//          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('H'.$cel, $row['DEPARTAMENTO']);

          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('G'.$cel, $row['PRO_NOMBRE']);         
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('H'.$cel, $row['MONTO_MAXIMO']);          
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('I'.$cel, $row['CUOTAS']);        
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('J'.$cel, $row['MONTO CREDITO']);  
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('K'.$cel, $row['CUOTAS OTORGADAS']);  
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('L'.$cel, $row['PERIODO_GRACIA']);      
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('M'.$cel, $row['ESTADO']);          

          if (($row['ULTIMO_PAGO'] != null) || ($row['ULTIMO_PAGO'] != 0))
          {
            $datetime= $row['ULTIMO_PAGO'];
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('N'.$cel, PHPExcel_Shared_Date::PHPToExcel($datetime));
            $objPHPExcel->getActiveSheet()->getStyle('N'.$cel)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME);
          }
          else
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('N'.$cel, '');
          
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('O'.$cel, $row['MOROSIDAD CAPITAL']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('P'.$cel, $row['CUOTAS ADEUDADAS']);         
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('Q'.$cel, $row['CUOTAS VENCIDAS']);         

          if (($row['FECHA_ULTIMA_ACTUALIZACION'] != null) || ($row['FECHA_ULTIMA_ACTUALIZACIONFECHA_ULTIMA_ACTUALIZACION'] != 0))
          {
            $datetime2= $row['FECHA_ULTIMA_ACTUALIZACION'];
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('R'.$cel, PHPExcel_Shared_Date::PHPToExcel($datetime2));
            $objPHPExcel->getActiveSheet()->getStyle('R'.$cel)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME);
          }
          else
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('R'.$cel, '');
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('S'.$cel, $row['ULTIMA_NOVEDAD']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('T'.$cel, $row['DEUDA CAPITAL']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('U'.$cel, $row['SOL_PER_MAIL']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('V'.$cel, $row['SOL_PER_TEL_FIJO']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('W'.$cel, $row['SOL_PER_TEL_CEL']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('X'.$cel, $row['CARACTER_GARANTE']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('Y'.$cel, $row['NOMBRE_GARANTE']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('Z'.$cel, $row['DOMICILIO_GARANTE']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AA'.$cel, $row['GARANTES']);
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AB'.$cel, $row['TELEFONOS']);        
          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AC'.$cel, $row['ATM']);  
          
          if (($row['FECHA_EFECTIVIZACION'] != null) || ($row['FECHA_EFECTIVIZACION'] != 0))
          {
            $datetime2= $row['FECHA_EFECTIVIZACION'];
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AD'.$cel, PHPExcel_Shared_Date::PHPToExcel($datetime2));
            $objPHPExcel->getActiveSheet()->getStyle('AD'.$cel)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME);
          }
          else
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AD'.$cel, '');

//          $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AC'.$cel, $row['FECHA_EFECTIVIZACION']);  
	        $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AE'.$cel, $row['LOC_ID']);  
	        $objPHPExcel->setActiveSheetIndex(0)->setCellValue('AF'.$cel, $row['LOC_NOMBRE']);  

          $cel++;
        }

        header("Content-type: application/vnd.ms-excel; charset=UTF-8" );
        header('Content-Disposition: attachment;filename="'.$anio.'-'.$mes.'-REPORTE_PROCURACION.xls"');
        header('Cache-Control: max-age=0');
         
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save('php://output');


  
?>
