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

  $conn = sqlsrv_connect($server,$concetinfo);

  $sql = "SELECT count(C.ID_SOLICITUD) as Cantidad
          FROM FBC_CUOTA CC INNER JOIN FBC_CREDITO C ON C.CRE_ID = CC.ID_CREDITO 
                      INNER JOIN FBC_PAGO PA ON PA.PAG_ID = CC.ID_PAGO 
                INNER JOIN FBC_SOLICITUD S ON S.SOL_ID = C.ID_SOLICITUD 
                      WHERE MONTH(PA.PAG_FECHA_PAGO) = $mes AND YEAR(PA.PAG_FECHA_PAGO) = $año
                AND S.ID_PROGRAMA_SOLICITADO = 43120";

$stmt=sqlsrv_query($conn,$sql);

echo '<div class="text-center" >';
echo '    <h4>REPORTE DE CONAMI</b></h4>';
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
                                       <button type="button" class="btn btn-light" id="botons"><a href ="javascript:openPage('.$mes.','.$año.')">Descargar Excel <span class="glyphicon glyphicon-download-alt"></span></button>
                                    </div>
                            </fieldset> 
                          </form>
                        </div>
                     </div> 
                </div>


<script language="javascript" type="text/javascript">

openPage = function($mes,$año) {
location.href = "excel_reporteconami.php?mes="+$mes+"&año="+$año;
 
}
</script>'
;

?>
