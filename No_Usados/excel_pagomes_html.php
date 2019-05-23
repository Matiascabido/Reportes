<?php
  include 'conn.php';
  $mes = $_POST['fecham'];
  $año = $_POST['fechaa'];



echo '                              <div class="col-md-6 col-md-offset-4">
                                       <div class="col-md-8 col-md-offset-2">
                                       <button type="button" class="btn  id="botons"><a href ="javascript:openPage('.$mes.','.$año.')">Descargar Excel <span class="glyphicon glyphicon-download-alt"></span></button>
                                       </div>
                                    </div>
                            </fieldset> 
                          </form>
                        </div>
                     </div> 
                </div>


<script language="javascript" type="text/javascript">

openPage = function($mes,$año) {
location.href = "excel_pagomes.php?mes="+$mes+"&año="+$año;
 
}
</script>'
;

?>
