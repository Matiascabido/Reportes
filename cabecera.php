<nav class="navbar navbar-inverse" style="background-color: rgb(8,146,208);">
  <div class="container-fluid">
     <div class="navbar-header">
      <a class="navbar-text"><b>Usuario: <b><?php echo strtoupper($_SESSION['currUsuario']); ?></b></b></a>
  </div>
  <div class="navbar-header">
   <div class="dropdown">
    <button type="button" class="btn btn-primary dropdown-toggle" data-toggle="dropdown">Reportes</button>
    <div class="dropdown-menu">
      <a class="dropdown-item" href="reportesG.php">Generales</a>
      <a class="dropdown-item" href="reportesXusuario.php">Usuario</a>
      <a class="dropdown-item" href="reportesXfecha.php">Fecha</a>
    </div>
      <input class="btn btn-primary" type="button" value="Cerrar Sesión" onclick="cerrarSession()">
      </div>
  </div>
  </div>
</nav>
