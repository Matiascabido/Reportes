<?php 

session_start();

ini_set('mssql.timeout',1000);
set_time_limit(1000);

$server = "fbcoprd.database.windows.net";
$user = "adminfbco";
$pwd="Fundacion#123";
$dba="GestionCreditosFBCO";
$concetinfo=array("Database" =>$dba , "UID" =>$user, "PWD"=>$pwd, "CharacterSet" => "UTF-8");
$conn = sqlsrv_connect($server,$concetinfo);


function ultimoId($conn, $tabla,$campo)
{
	$res =0;
	$sql = "SELECT ISNULL(MAX(".$campo."),0) +1 FROM ".$tabla;
	//return $sql;
	$r = sqlsrv_query($conn,$sql);
	$row = sqlsrv_fetch_array($r);
	return $row[0];
}

?>