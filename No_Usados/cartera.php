<?php
  include 'conn.php';
  $mes = $_POST['fecham'];
  $año = $_POST['fechaa'];
  $server = "fbcoprd.database.windows.net";
  $user = "adminfbco";
  $pwd="Fundacion#123";
  $dba="GestionCreditosFBCO";

  $concetinfo=array("Database" =>$dba , "UID" =>$user, "PWD"=>$pwd, "CharacterSet" => "UTF-8");
  $conn = sqlsrv_connect($server,$concetinfo);

$sql = " SELECT 'FONCAP' as Programas , count(*) as Cantidad 
         FROM FBC_CREDITO C INNER JOIN FBC_SOLICITUD S ON S.SOL_ID = C.ID_SOLICITUD 
         INNER JOIN FBC_PROGRAMA P ON P.PRO_ID = S.ID_PROGRAMA_SOLICITADO 
         INNER JOIN FBC_PERSONA PE ON PE.PER_ID = S.ID_TITULAR 
         INNER JOIN FBC_LOCALIDAD L ON L.LOC_ID = S.SOL_ID_LOCALIDAD 
         INNER JOIN FBC_DEPARTAMENTO D ON D.DTO_ID = L.ID_DEPARTAMENTO 
         INNER JOIN FBC_EMPRENDIMIENTO E ON E.EMP_ID = S.ID_EMPRENDIMIENTO 
         WHERE Year(C.CRE_FECHA_EFECTIVIZACION) = $año And Month(C.CRE_FECHA_EFECTIVIZACION) = $mes
         AND S.ID_PROGRAMA_SOLICITADO = 43108
         union all
          SELECT 'NO FONCAP' as Programas , count(*) as Cantidad 
         FROM FBC_CREDITO C INNER JOIN FBC_SOLICITUD S ON S.SOL_ID = C.ID_SOLICITUD 
         INNER JOIN FBC_PROGRAMA P ON P.PRO_ID = S.ID_PROGRAMA_SOLICITADO 
         INNER JOIN FBC_PERSONA PE ON PE.PER_ID = S.ID_TITULAR 
         INNER JOIN FBC_LOCALIDAD L ON L.LOC_ID = S.SOL_ID_LOCALIDAD 
         INNER JOIN FBC_DEPARTAMENTO D ON D.DTO_ID = L.ID_DEPARTAMENTO 
         INNER JOIN FBC_EMPRENDIMIENTO E ON E.EMP_ID = S.ID_EMPRENDIMIENTO 
         WHERE Year(C.CRE_FECHA_EFECTIVIZACION) = $año And Month(C.CRE_FECHA_EFECTIVIZACION) = $mes 
         AND S.ID_PROGRAMA_SOLICITADO <> 43108";




echo '
         <div class="col-md-6 col-md-offset-4">
                    <div>
                        <div class="panel-body">
                             <form role="form" >
                                <fieldset>';
echo '                              <div class="col-md-6 col-md-offset-1">';
//echo $sql;

$stmt=sqlsrv_query($conn,$sql);

echo '<div class="text-center" >';
echo '    <h4>CARTERA INCORPORADA</b></h4>';
echo '    <h6>(Generar el Excel puede demorar varios minutos)</b></h6>';
echo '</div>';

echo '<table class="table" style="text-align: center;" border ="1">';
echo '  <tr style="text-align: center; background:#0099CC">';
echo '    <th style="text-align: center;">Programas </th>';
echo '    <th style="text-align: center;">Cantidad </th>';
echo "  </tr>";
while ($row = sqlsrv_fetch_array($stmt))
{
  echo "<tr>";
  echo "<td>".$row['Programas']."</td>";
  echo "<td>".$row['Cantidad']."</td>";
  echo "</tr>";
}
echo "</table>";
echo "<br>";
echo '</div>';
echo '                              <div class="col-md-6 col-md-offset-2">
                                       <button type="button" class="btn "><a href ="javascript:openPage('.$mes.','.$año.')">Descargar Excel <span class="glyphicon glyphicon-download-alt"></span></button>

                                    </div>
                            </fieldset> 
                          </form>
                        </div>
                     </div> 
                </div>


<script language="javascript" type="text/javascript">
openPage = function($mes,$año) {
location.href = "Excel/activados_mes.php?mes="+$mes+"&año="+$año;

}
</script>'
;

?>
