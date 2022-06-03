<?php
$dirname = dirname(__DIR__);
include_once  $dirname.'/common/DataServices.php';

class usuariosBD {
    //Otener todos los usuarios
    public function obtTodosUsuariosDB($limit = ""){
        $ds = new DataServices();
        $param[0] = "";
        $param[1] = "";
        $query = array();

        // if($idRol != ""){
        //     $query[] = " A.idRol IN( $idRol )";//JGP cambio por IN para permtir separado por comas
        // }

        // if($idAbogado > 0){
        //     // echo "id abnogado ".$idAbogado;
        //     $query[] = " A.idUsuario=$idAbogado ";
        // }
        // //En caso de llevar filtro
        // if(count($query) > 0){
        //   $wordWhere = " WHERE ";
        //   $setWhere = implode(" AND ", $query);
        //   // echo $setWhere;
        //   $param[0] = $wordWhere.$setWhere;
        // }

        // //Jair 13/4/2022 Ordenar
        // if($order != ""){
        //     $param[1] = " ORDER BY $order ";
        // }

        if($limit != ""){
            $param[1] = " LIMIT $limit ";
        }

        $result = $ds->Execute("obtTodosUsuariosDB", $param);
        $ds->CloseConnection();
        return $result;
    }

    public function updateUsuarioDB($param){
        $ds = new DataServices();
        $result = $ds->Execute("updateUsuarioDB", $param, false, true);
        $ds->CloseConnection();
        return $result;
    }

    public function insertUsuarioDB($param){
        $ds = new DataServices();
        $result = $ds->Execute("insertUsuarioDB", $param, true);
        return $result;
    }

    public function insertEmpleadoDB($param){
        $ds = new DataServices();
        $result = $ds->Execute("insertEmpleadoDB", $param, true);
        return $result;
    }
}