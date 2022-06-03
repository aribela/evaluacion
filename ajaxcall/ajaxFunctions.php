<?php 
$dirname = dirname(__DIR__);
include_once $dirname.'/brules/usuariosObj.php';
$function= $_GET['funct'];
switch ($function){
    case "funcionEjemplo":
        funcionEjemplo($_GET['callback'], $_GET['id'],$_GET['valor']);
    break;

    case "verPrimeros":
        verPrimeros();
    break;

    case "actualizaSueldo":
        actualizaSueldo();
    break;

    case "guardar":
        guardar();
    break;

}

function funcionEjemplo($callback, $id, $valor){
    $gamaModelosObj = new gamaModelosObj();
    $actualizacionesObj = new catActualizacionesObj();
    $actualizacionesObj->updActualizacion("gama_modelos");
    $res = $gamaModelosObj->ActCampoGamaModelo('activo', $valor, $id);
    $return_arr = array("success"=>true, "res"=>$res);
    echo $callback . '(' . json_encode($return_arr) . ');';
}

function verPrimeros(){
    $callback = $_GET['callback'];
    $limit = $_GET['limit'];

    $usuariosObj = new usuariosObj();
    $usuarios = $usuariosObj->obtTodosUsuarios(true, $limit);

    $html = '';

    //html table
    $html .= '<table class="table table-bordered table-striped">';
    $html .= '<thead>';
    $html .= '<tr>';
    $html .= '<th>ID</th>';
    $html .= '<th>Nombre</th>';
    $html .= '<th>Apellido Paterno</th>';
    $html .= '<th>Apellido Materno</th>';
    $html .= '<th>Sueldo</th>';
    $html .= '<th></th>';
    $html .= '</tr>';
    $html .= '</thead>';
    $html .= '<tbody>';
    foreach ($usuarios as $key => $value) {
        $html .= '<tr>';
        $html .= '<td>'.$value->userId.'</td>';
        $html .= '<td>'.$value->Nombres.'</td>';
        $html .= '<td>'.$value->Paterno.'</td>';
        $html .= '<td>'.$value->Materno.'</td>';
        $html .= '<td>'.$value->Sueldo.'</td>';
        $html .= '<td>
        <input type="text" name="sueldo_'.$value->userId.'", id="sueldo_'.$value->userId.'" value="'.$value->Sueldo.'" class="form-control" onchange="actualizaSueldo('.$value->userId.', this.value)">
        </td>';
       
        $html .= '</tr>';
    }
    $html .= '</tbody>';
    $html .= '</table>';


    $return_arr = array("success"=>true, "html"=>$html);
    echo $callback . '(' . json_encode($return_arr) . ');';
}

function actualizaSueldo(){
    $callback = $_GET['callback'];
    $userId = $_GET['userId'];
    $sueldo = $_GET['sueldo'];

    $usuariosObj = new usuariosObj();
    $usuariosObj->ActualizarUsuario('sueldo', $sueldo, $userId);

    $return_arr = array("success"=>true);
    echo $callback . '(' . json_encode($return_arr) . ');';
}

function guardar(){
    $nombre = $_POST['nombre'];
    $Paterno = $_POST['Paterno'];
    $Materno = $_POST['Materno'];
    $sueldo = $_POST['Sueldo'];
    $usuariosObj = new usuariosObj();
    $usuariosObj->insertUsuario($nombre, $Paterno, $Materno, $sueldo);

    $arr = array("success"=>true, "post"=>$_POST);

  echo json_encode($arr);
}
