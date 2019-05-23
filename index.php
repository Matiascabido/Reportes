<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <meta http-equiv="Expires" content="0">
    <meta http-equiv="Last-Modified" content="0">
    <meta http-equiv="Cache-Control" content="no-cache, mustrevalidate">
    <meta http-equiv="Pragma" content="no-cache">
    <meta http-equiv="x-ua-compatible" content="ie=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="shortcut icon" href="">
    <link href="css/simple-sidebar.css" rel="stylesheet">
    <link href="//netdna.bootstrapcdn.com/bootstrap/3.0.3/css/bootstrap.min.css" rel="stylesheet"  id="bootstrap-css" />
    <title>FBCO | Sistema de Gestion de Creditos - Carga de Solicitud Reducida</title>
    <style >

      .loader {
        border: 16px solid #f3f3f3; /* Light grey */
        border-top: 16px solid #3498db; /* Blue */
        border-radius: 25%;
        width: 12px;
        height: 12px;
        animation: spin 2s linear infinite;
        }

        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
    </style>
    <?php include "head.php";?>
    <?php

function getDatetimeMonth()
{
    $tz_object = new DateTimeZone('Brazil/East');
    //date_default_timezone_set('Brazil/East');

    $datetime = new DateTime();
    $datetime->setTimezone($tz_object);
    $datetime->modify('-1 month');
    return $datetime->format('m');

}
function getDatetimeYear()
{
    $tz_object = new DateTimeZone('Brazil/East');
    //date_default_timezone_set('Brazil/East');

    $datetime = new DateTime();
    $datetime->setTimezone($tz_object);
    if (getDatetimeMonth() == 12) {
        $datetime->modify('-1 year');
        return $datetime->format('Y');
    }
    return $datetime->format('Y');

}

?>

</head>
    <div id="wrapper">
        <!-- Sidebar -->
        <div id="sidebar-wrapper">
            <ul class="sidebar-nav">
                <li class="sidebar-brand">
                    <a href="indes.php">
                        FBCO - REPORTES
                    </a>
                </li>
                <li>
                    <a href="index.php">- FONCAP</a>
                </li>
                <li>
                    <a href="indes.php">- OTROS</a>
                </li>
                <li>
                    <a href="indess.php">- CORTES</a>
                </li>
            </ul>
        </div>
        <!-- /#sidebar-wrapper -->
<body>
  <main class="content view-animate fade-up" id="content" role="main">
    <div class="row row justify-content-center mt-5 pt-5">
      <div class="col-lg-12 col-xs-12 breadcrumb">
        <section class="widget breadcrumb" widget="">
          <div>
          <div class="widget-body">
            <a href="#menu-toggle" style="background-color: black"; class="btn" id="menu-toggle"><span class="glyphicon glyphicon-tasks"></span></a>
             <ol class="breadcrumb">
              <li class="breadcrumb-item"><h3 style="text-align: center"> REPORTES FONCAP</h3></li>
                    <div class="col-md-6 col-md-offset-4">
                    <div>
                        <div class="panel-body">
                            <form role="form" >
                                <fieldset>
                                   <div class="col-md-6 col-md-offset-2 breadcrumb">
                                  <b><i class="  glyphicon glyphicon-calendar"></i></b>
                                     <input type="text" name="fecha" id="fechames" class="form-control ng-untouched ng-pristine ng-invalid" style="width: 20% !important; display: inline !important;" value="<?php echo getDatetimeMonth() ?>" placeholder="<?php echo getDatetimeMonth() ?> ">
                                     <input type="text" name="fecha" id="fechaAno" class="form-control ng-untouched ng-pristine ng-invalid" style="width: 25% !important; display: inline !important;"value="<?php echo getDatetimeYear() ?>" placeholder="<?php echo getDatetimeYear() ?>" >
                                   </div>
                                    <br>
                                    <br>

                                    <br>
                                    <div class="col-md-6 col-md-offset-1 breadcrumb">
                                     <button id="btnrecupero" type="button" class="btn btn-primary btn-block" onclick="Recupero();">Comportamiento Pago</button>
                                    </div>
                                    <br>
                                    <br>
                                    <br>

                                    <div class="col-md-6 col-md-offset-1 breadcrumb">
                                     <button  id="btncartera" type="button" class="btn btn-primary btn-block" onclick="Cartera();">Cartera Incorporada</button>
                                    </div>
                                    <br>
                                    <br>
                                    <br>
                                    <div class="col-md-6 col-md-offset-1 breadcrumb">
                                     <button type="button"  id="btncarteracomp" class="btn btn-primary btn-block" onclick="CarteraCompleta();">Cartera Completa</button>
                                    </div>
                                    <br>
                                    <br>
                                    <br>
                                   <div class="col-md-6 col-md-offset-1 breadcrumb">
                                     <button type="button"  id="btnpistola" class="btn btn-primary btn-block" onclick="Recupital();">Recupero Capital</button>
                                    </div>
                                    <br>
                                    <br>
                                    <br>
                                    <div class="col-md-6 col-md-offset-3 breadcrumb">
                                     <div id="loader" class="loader" style="display:none;">
                                     </div>
                                   </div>
                              </fieldset>
                          </form>
                        </div>
                      </div>
                     </div>
                     </ol>
                   </div>
                  <div id="Recupero">
                 </div>
              </ol>
            </div>
           </div>
        </section>
      </div>
    </div>
  </main>
</div>
    </div>
    <!-- /#wrapper -->
    
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.4.0/jquery.min.js"></script>
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.4.0/js/bootstrap.min.js"></script>
    <script src="js/codigo.js"></script>
    <!-- Menu Toggle Script -->
    <script>
    $("#menu-toggle").click(function(e) {
        e.preventDefault();
        $("#wrapper").toggleClass("toggled");
    });
    </script>
<script language="javascript" type="text/javascript">
openPage = function($mes,$año) {

location.href = "excel.php?mes="+$mes+"&año="+$año;

}
</script>
</body>
</html>
