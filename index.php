<?php 

?>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document</title>
    <link href="js/boostrap-select/bootstrap-select.min.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="js/jquery-1.10.2.min.js"></script>
    <script type="text/javascript" src="js/boostrap-select/bootstrap-select.min.js"></script>
    <script type="text/javascript" src="js/functionsGlobals.js?upd=<?php echo time(); ?>"></script> 
    <script type="text/javascript" src="js/functions.js?upd=<?php echo time(); ?>"></script>
</head>
<body>
    <button onclick="verPrimeros(10)">Ver primeros 10</button>
    <button onclick="verPrimeros('')">Ver todos</button>
    <button onclick="verNuevo('')">Nuevo</button>
    <div id="contenido">

    </div>
</body>
</html>