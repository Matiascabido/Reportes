<?php
  include 'conn.php';
  $mes = $_POST['fecham'];
  $año = $_POST['fechaa'];




echo '
         <div class="col-md-6 col-md-offset-4">
                    <div>
                        <div class="panel-body">
                             <form role="form" >
                                <fieldset>';
echo '                              <div class="col-md-6 col-md-offset-1">';

  $mes = $_POST['fecham'];  
  $año = $_POST['fechaa'];
  $conn = sqlsrv_connect($server,$concetinfo);

  $sql = "SELECT count(*) as Cantidad
          /*
           C.ID_SOLICITUD as [Nº Operacion], P.PER_NOMBRE + ' ' + P.PER_APELLIDO  as [Nombre y Apellido del Cliente], P.PER_FEC_NAC as [Fecha de Nacimiento],  P.PER_TIPO_DOC as [Tipo de Documento], 
          P.PER_NUM_DOC as [Numero de Documento], concat(P.PER_CALLE,' ', P.PER_NUM_CALLE,' ', P.PER_PISO,' ',P.PER_DPTO, ' - ', 
          L.loc_nombre,' (' ,de.dto_nombre,')') as [Domicilio], P.PER_SEXO as [Genero], 
          C.CRE_FECHA_EFECTIVIZACION as [Fecha de Inicio de Microcredito],  
           C.CRE_MONTO_OTORGADO as [Suma Asegurada Inicial/Monto del Microcredito], (SELECT MAX(CU1.cuo_vencimiento_1) FROM FBC_CUOTA CU1 WHERE CU1.ID_CREDITO = C.CRE_ID) as [Fecha de Finalizacion del Microcredito],  
          isnull((C.CRE_MONTO_OTORGADO - isnull((SELECT SUM(CU.CUO_MONTO_CAPITAL)         FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS NOT NULL),0)),0) AS [Saldo Deudor de Capital al dia del reporte] ,  
          (select count(*) from FBC_CUOTA CU1 WHERE CU1.ID_CREDITO = C.CRE_ID AND CU1.ID_PAGO IS NULL AND CU1.cuo_vencimiento_1 < getdate()) as [CUOTAS ATRASADAS] 
          */
          FROM FBC_CREDITO C INNER JOIN FBC_SOLICITUD S ON S.SOL_ID = C.ID_SOLICITUD INNER JOIN FBC_PERSONA P ON P.PER_ID = S.ID_TITULAR          
          LEFT OUTER JOIN FBC_LOCALIDAD L ON L.LOC_ID = S.SOL_ID_LOCALIDAD         
          LEFT OUTER JOIN FBC_DEPARTAMENTO DE ON DE.DTO_ID = L.ID_DEPARTAMENTO         
          WHERE C.CRE_FECHA_EFECTIVIZACION > DATEFROMPARTS (2016,6,1)         
          and isnull((C.CRE_MONTO_OTORGADO - isnull((SELECT SUM(CU.CUO_MONTO_CAPITAL) FROM FBC_CUOTA CU  WHERE CU.ID_CREDITO = C.CRE_ID AND CU.ID_PAGO IS NOT NULL),0)),0) > 0 
          AND C.CRE_SEGURO_INFORMADO IS NULL";
$stmt=sqlsrv_query($conn,$sql);

echo '<div class="text-center" >';
echo '    <h4>REPORTE DE SEGURO</b></h4>';
//echo '    <h6>(Generar el Excel puede demorar varios minutos)</b></h6>';
echo '</div>';
echo '<table class="table" style="text-align: center;" border ="1">';
echo '  <tr style="text-align: center; background:#0099CC">';
echo '    <th style="text-align: center;">Cantidad </th>';
echo "  </tr>";
while ($row = sqlsrv_fetch_array($stmt))
{
  echo "<tr>";
  echo "<td>".$row['Cantidad']."</td>";
  echo "</tr>";
}
echo "</table>";

echo '<div class="text-center" >';
echo '    <h6>(Generar el Excel puede demorar)</b></h6>';
echo '</div>';


echo '                              <div class="col-md-6 col-md-offset-2">
                                       <button type="button" class="btn btn-light"  id="botons"><a href ="javascript:openPage('.$mes.','.$año.')">Descargar Excel <span class="glyphicon glyphicon-download-alt"></span></button>
                                    </div>
                            </fieldset> 
                          </form>
                        </div>
                     </div> 
                </div>


<script language="javascript" type="text/javascript">

openPage = function($mes,$año) {
location.href = "excel_reporteseguro.php?mes="+$mes+"&año="+$año;
 
}
</script>'
;

?>
