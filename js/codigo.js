function pulsar(e) {
  if (e.keyCode === 13 && !e.shiftKey) {
    filtrarCodigo();
  }
}
//FUNCIONES PARA LA PAG SECUNDARIA INDESS.PHP(PAGINA SECUNDARIA)
function Corte() {
  $("#Recupero").html("");
  $("#loader").show();
  $("#btncorte").attr("disabled", "true");
  $.ajax({
    url: "href_corte.php",
    data: { fechad:$("#fechadia").val(),fecham: $("#fechames").val(), fechaa: $("#fechaAno").val(),fechad2:$("#fechadia2").val(),fecham2: $("#fechames2").val(), fechaa2: $("#fechaAno2").val() },
    dataType: "text",
    type: "POST",
    success: function(data) {
      $("#Recupero")
        .fadeIn(1000)
        .html(data);
      $("#loader").fadeOut(700);
      $("#btncorte").removeAttr("disabled");
    }
  });
}

//FUNCIONES PARA LA PAG SECUNDARIA INDES.PHP(PAGINA SECUNDARIA)

function legajoPA() {
  $("#Recupero").html("");
  $("#loader").show();
  $("#btncarteracomp").attr("disabled", "true");
  $("#btncartera").attr("disabled", "true");
  $("#btnrecupero").attr("disabled", "true");
  $("#btnpistola").attr("disabled", "true");
  $("#btnLegajo").attr("disabled", "true");
  $.ajax({
    url: "Href/href_solicitudpa.php",
    data: { fecham: $("#fechames").val(), fechaa: $("#fechaAno").val() },
    dataType: "text",
    type: "POST",
    success: function(data) {
      $("#Recupero")
        .fadeIn(1000)
        .html(data);
      $("#loader").fadeOut(700);
      $("#btnLegajo").removeAttr("disabled");
      $("#btncarteracomp").removeAttr("disabled");
      $("#btncartera").removeAttr("disabled");
      $("#btnrecupero").removeAttr("disabled");
      $("#btnpistola").removeAttr("disabled");
    }
  });
}

function CapitalMes() {
  $("#Recupero").html("");
  $("#loader").show();
  $("#btncarteracomp").attr("disabled", "true");
  $("#btncartera").attr("disabled", "true");
  $("#btnrecupero").attr("disabled", "true");
  $("#btnpistola").attr("disabled", "true");
  $("#btnLegajo").attr("disabled", "true");
  $.ajax({
    url: "Href/href_capitalmes.php",
    data: { fecham: $("#fechames").val(), fechaa: $("#fechaAno").val() },
    dataType: "text",
    type: "POST",
    success: function(data) {
      $("#Recupero")
        .fadeIn(1000)
        .html(data);
      $("#loader").fadeOut(700);
      $("#btncarteracomp").removeAttr("disabled");
      $("#btnLegajo").removeAttr("disabled");
      $("#btncartera").removeAttr("disabled");
      $("#btnrecupero").removeAttr("disabled");
      $("#btnpistola").removeAttr("disabled");
    }
  });
}

// function Reporte6Meses() {
//   $("#Recupero").html("");
//   $("#loader").show();
//   $("#btncarteracomp").attr("disabled", "true");
//   $("#btncartera").attr("disabled", "true");
//   $("#btnrecupero").attr("disabled", "true");
//   $("#btnpistola").attr("disabled", "true");
//   $("#btnLegajo").attr("disabled", "true");
//   $.ajax({
//     url: "Href/href_6meses.php",
//     data: { fecham: $("#fechames").val(), fechaa: $("#fechaAno").val() },
//     dataType: "text",
//     type: "POST",
//     success: function(data) {
//       $("#Recupero")
//         .fadeIn(1000)
//         .html(data);
//       $("#loader").fadeOut(700);
//       $("#btncarteracomp").removeAttr("disabled");
//       $("#btnLegajo").removeAttr("disabled");
//       $("#btncartera").removeAttr("disabled");
//       $("#btnrecupero").removeAttr("disabled");
//       $("#btnpistola").removeAttr("disabled");
//     }
//   });
// }

function ReporteProcuracion() {
  $("#Recupero").html("");
  $("#loader").show();
  $("#btncarteracomp").attr("disabled", "true");
  $("#btncartera").attr("disabled", "true");
  $("#btnrecupero").attr("disabled", "true");
  $("#btnpistola").attr("disabled", "true");
  $("#btnLegajo").attr("disabled", "true");
  $.ajax({
    url: "Href/href_reporteprocuracion.php",
    data: { fecham: $("#fechames").val(), fechaa: $("#fechaAno").val() },
    dataType: "text",
    type: "POST",
    success: function(data) {
      $("#Recupero")
        .fadeIn(1000)
        .html(data);
      $("#loader").fadeOut(700);
      $("#btncarteracomp").removeAttr("disabled");
      $("#btnLegajo").removeAttr("disabled");
      $("#btncartera").removeAttr("disabled");
      $("#btnrecupero").removeAttr("disabled");
      $("#btnpistola").removeAttr("disabled");
    }
  });
}

function ReporteConami() {
  $("#Recupero").html("");
  $("#loader").show();
  $("#btncarteracomp").attr("disabled", "true");
  $("#btncartera").attr("disabled", "true");
  $("#btnrecupero").attr("disabled", "true");
  $("#btnpistola").attr("disabled", "true");
  $("#btnLegajo").attr("disabled", "true");
  $.ajax({
    url: "Href/href_reporteconami.php",
    data: { fecham: $("#fechames").val(), fechaa: $("#fechaAno").val() },
    dataType: "text",
    type: "POST",
    success: function(data) {
      $("#Recupero")
        .fadeIn(1000)
        .html(data);
      $("#loader").fadeOut(700);
      $("#btncarteracomp").removeAttr("disabled");
      $("#btnLegajo").removeAttr("disabled");
      $("#btncartera").removeAttr("disabled");
      $("#btnrecupero").removeAttr("disabled");
      $("#btnpistola").removeAttr("disabled");
    }
  });
}

function ReporteSeguro() {
  $("#Recupero").html("");
  $("#loader").show();
  $("#btncarteracomp").attr("disabled", "true");
  $("#btncartera").attr("disabled", "true");
  $("#btnrecupero").attr("disabled", "true");
  $("#btnpistola").attr("disabled", "true");
  $("#btnLegajo").attr("disabled", "true");
  $.ajax({
    url: "Href/href_reporteseguro.php",
    data: { fecham: $("#fechames").val(), fechaa: $("#fechaAno").val() },
    dataType: "text",
    type: "POST",
    success: function(data) {
      $("#Recupero")
        .fadeIn(1000)
        .html(data);
      $("#loader").fadeOut(700);
      $("#btncarteracomp").removeAttr("disabled");
      $("#btncartera").removeAttr("disabled");
      $("#btnLegajo").removeAttr("disabled");
      $("#btnrecupero").removeAttr("disabled");
      $("#btnpistola").removeAttr("disabled");
    }
  });
}

//FUNCIONES PARA LA PAGINA INDEX.PHP (PAGINA PRINCIPAL)

function Recupital() {
  $("#Recupero").html("");
  $("#loader").show();
  $("#btncarteracomp").attr("disabled", "true");
  $("#btncartera").attr("disabled", "true");
  $("#btnrecupero").attr("disabled", "true");
  $("#btnpistola").attr("disabled", "true");
  $.ajax({
    url: "Href/href_recupital.php",
    data: { fecham: $("#fechames").val(), fechaa: $("#fechaAno").val() },
    dataType: "text",
    type: "POST",
    success: function(data) {
      $("#Recupero")
        .fadeIn(1000)
        .html(data);
      $("#loader").fadeOut(700);
      $("#btncarteracomp").removeAttr("disabled");
      $("#btncartera").removeAttr("disabled");
      $("#btnrecupero").removeAttr("disabled");
      $("#btnpistola").removeAttr("disabled");
    }
  });
}

function CarteraCompleta() {
  $("#Recupero").html("");
  $("#loader").show();
  $("#btncarteracomp").attr("disabled", "true");
  $("#btncartera").attr("disabled", "true");
  $("#btnrecupero").attr("disabled", "true");
  $("#btnpistola").attr("disabled", "true");
  $.ajax({
    url: "Href/href_carteraCompleta.php",
    data: { fecham: $("#fechames").val(), fechaa: $("#fechaAno").val() },
    dataType: "text",
    type: "POST",
    success: function(data) {
      $("#Recupero")
        .fadeIn(1000)
        .html(data);
      $("#loader").fadeOut(700);
      $("#btncarteracomp").removeAttr("disabled");
      $("#btncartera").removeAttr("disabled");
      $("#btnrecupero").removeAttr("disabled");
      $("#btnpistola").removeAttr("disabled");
    }
  });
}

function Cartera() {
  $("#Recupero").html("");
  $("#loader").show();
  $("#btncarteracomp").attr("disabled", "true");
  $("#btncartera").attr("disabled", "true");
  $("#btnrecupero").attr("disabled", "true");
  $("#btnpistola").attr("disabled", "true");
  $.ajax({
    url: "Href/href_cartera.php",
    data: { fecham: $("#fechames").val(), fechaa: $("#fechaAno").val() },
    dataType: "text",
    type: "POST",
    success: function(data) {
      $("#Recupero")
        .fadeIn(1000)
        .html(data);
      $("#loader").fadeOut(700);
      $("#btncarteracomp").removeAttr("disabled");
      $("#btncartera").removeAttr("disabled");
      $("#btnrecupero").removeAttr("disabled");
      $("#btnpistola").removeAttr("disabled");
    }
  });
}

function Recupero() {
  $("#Recupero").html("");
  $("#loader").show();
  $("#btncarteracomp").attr("disabled", "true");
  $("#btncartera").attr("disabled", "true");
  $("#btnrecupero").attr("disabled", "true");
  $("#btnpistola").attr("disabled", "true");
  $.ajax({
    url: "Href/href_recupero.php",
    data: { fecham: $("#fechames").val(), fechaa: $("#fechaAno").val() },
    dataType: "text",
    type: "POST",
    success: function(data) {
      $("#Recupero")
        .fadeIn(1000)
        .html(data);
      $("#loader").fadeOut(1000);
      $("#btncarteracomp").removeAttr("disabled");
      $("#btncartera").removeAttr("disabled");
      $("#btnrecupero").removeAttr("disabled");
      $("#btnpistola").removeAttr("disabled");
    }
  });
}

