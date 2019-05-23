<?php

  $mes = $_POST['fecham'];
  $año = $_POST['fechaa'];




echo '
         <div class="col-md-6 col-md-offset-4">
                    <div>
                        <div class="panel-body">
                             <form role="form" >
                                <fieldset>';
echo '                              <div class="col-md-6 col-md-offset-1">';

echo '<div class="text-center" >';

echo '    <h6>(Generar el Excel puede demorar hasta 10 minutos)</b></h6>';
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
location.href = "excel_reporteprocuracion.php?mes="+$mes+"&año="+$año;
 
}
</script>'
;

?>
