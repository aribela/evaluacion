<?php 
$dirname = dirname(__DIR__);
include_once  '../database/usuariosBD.php';
include_once  '../database/datosBD.php';
include_once  '../brules/configuracionesGridObj.php';

// include_once  '../brules/generalConsultObj.php';

class usuariosObj extends configuracionesGridObj{
    private $_userId = 0;
    private $_Login = "";
    private $_Nombres = "";
    private $_Paterno = ""; 
    private $_Materno = "";

    //Obtener todos los usuarios
    public function obtTodosUsuarios($opcObj=true, $limit = ""){
        $array = array();
        $ds = new usuariosBD();
        $datosBD = new datosBD();
        $result = $ds->obtTodosUsuariosDB($limit);

        if($opcObj==true){
           $array = (array)$datosBD->arrDatosObj($result);
            // $array = (array)arrDatosObj($result);
        }else{
           $array = $datosBD->arrDatosObj($result);
           // $array = arrDatosObj($result);
        }
        return $array;
    }

    public function ActualizarUsuario($campo, $valor, $id){
        $param[0] = $campo;
        $param[1] = $valor;
        $param[2] = $id;

        $objDB = new usuariosBD();
        $resAct = $objDB->updateUsuarioDB($param);
        return $resAct;
    }

    public function insertUsuario($nombre, $Paterno, $Materno, $sueldo){
        
        $param[0] = "login_".$nombre;
        $param[1] = $nombre;
        $param[2] = $Paterno;
        $param[3] = $Materno;
        $objDB = new usuariosBD();
        $this->_userId = $objDB->insertUsuarioDB($param);
        $param2[0] = $this->_userId;
        $param2[1] = $sueldo;
        $empleadoId = $objDB->insertEmpleadoDB($param2);
    }
}