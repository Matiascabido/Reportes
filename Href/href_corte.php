<?php
  include 'conn.php';
  $mes = $_POST['fecham'];
  $año = $_POST['fechaa'];
  $dia = $_POST['fechad'];
  $mes2 = $_POST['fecham2'];
  $año2 = $_POST['fechaa2'];
  $dia2= $_POST['fechad2'];


  $sql = "";

echo '                              <div class="col-md-6 col-md-offset-2">
                                       <button type="button" class="btn btn-light" id="butonx"><a href ="javascript:openPage('.$dia.','.$mes.','.$año.','.$dia2.','.$mes2.','.$año2.')">Descargar Excel <span class="glyphicon glyphicon-download-alt"></span></button>
                                    </div>
                            </fieldset> 
                          </form>
                        </div>
                     </div> 
                </div>


<script language="javascript" type="text/javascript">

openPage = function($mes,$año) {
location.href = "excel_corte.php?dia="+$dia+"&mes="+$mes+"&año="+$año+"&dia2="+$dia2+"&mes2="+$mes2+"&año2="+$año2;
}
</script>'
;

?>
