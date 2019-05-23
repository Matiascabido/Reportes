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




$sql = " select cast(sum(capital) as decimal(19,2)) as Capital, cast(sum(interes) as decimal(19,2)) as Interes from (
SELECT C.ID_SOLICITUD [CREDITO], CC.CUO_NUMERO [CUOTA] ,CONCAT(S.SOL_PER_APELLIDO, ' ', S.SOL_PER_NOMBRE) AS [NOMBRE], 
CAST(CC.CUO_MONTO_CAPITAL as DECIMAL(19,4)) AS [CAPITAL], SUM(CASE WHEN (DATEDIFF(D,CC.CUO_VENCIMIENTO_1, PA.PAG_FECHA_PAGO)) <= 0 
THEN CUO_MONTO_INTERES_FINANCIERO_1 ELSE CASE WHEN (DATEDIFF(D,CC.CUO_VENCIMIENTO_2, PA.PAG_FECHA_PAGO)) <= 0  
THEN CUO_MONTO_INTERES_FINANCIERO_2 ELSE CASE WHEN (DATEDIFF(D,CC.CUO_VENCIMIENTO_3, PA.PAG_FECHA_PAGO)) <= 0  
THEN CUO_MONTO_INTERES_FINANCIERO_3 ELSE CUO_MONTO_INTERES_FINANCIERO_1+CUO_MONTO_INTERES_PUNITORIO END END END) AS INTERES 
FROM FBC_CUOTA CC INNER JOIN FBC_CREDITO C ON C.CRE_ID = CC.ID_CREDITO INNER JOIN FBC_PAGO PA ON PA.PAG_ID = CC.ID_PAGO 
INNER JOIN FBC_SOLICITUD S ON S.SOL_ID = C.ID_SOLICITUD 
WHERE Month(PA.PAG_FECHA_PAGO) = $mes And Year(PA.PAG_FECHA_PAGO) = $año 
GROUP BY C.ID_SOLICITUD, CC.CUO_NUMERO, 
CONCAT(S.SOL_PER_APELLIDO, ' ', S.SOL_PER_NOMBRE), CC.CUO_MONTO_CAPITAL ) as X";

$stmt=sqlsrv_query($conn,$sql);

echo '<div style="text-align:center;">';
echo '    <h4>PAGOS DEL MES</b></h4>';
echo '    <h6>(Generar el Excel puede demorar varios minutos)</b></h6>';
echo '</div>';
echo '<table class="table" style="text-align: center;" border ="1">';
echo ' <thead class="bg-primary">';
echo '  <tr style="text-align: center; background:#0099CC">';
echo '    <th style="text-align: center;">Capital </th>';
echo '    <th style="text-align: center;">Interes </th>';
echo ' </tr> ' ;
echo '</thead>';
while ($row = sqlsrv_fetch_array($stmt))
{
  echo "<tr>";
  echo "<td> $ ".$row['Capital']."</td>";
  echo "<td> $ ".$row['Interes']."</td>";
  echo "</tr>";
}
echo "</table>";

echo '                              <div class="col-md-6 col-md-offset-2">
                                       <button type="button" class="btn  btn-light"  id="botons"><a href ="javascript:openPage('.$mes.','.$año.')">Descargar Excel <span class="glyphicon glyphicon-download-alt"></span></button>
                                    </div>
                            </fieldset> 
                          </form>
                        </div>
                     </div> 
                </div>


<script language="javascript" type="text/javascript">

openPage = function($mes,$año) {
location.href = "Excel/excel_capitalmes.php?mes="+$mes+"&año="+$año;
 
}
</script>'
;

?>
