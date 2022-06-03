<?php
$dirname = dirname(__DIR__);
include_once  $dirname.'/database/generalConsultBD.php';
include_once  $dirname.'/database/datosBD.php';
include_once  $dirname.'/brules/rolesObj.php';

class generalConsultObj {
  //Obtener coleccion de comunicadoss
    public function GetAllData($tabla, $activo=-1, $query = array(), $order = "", $leftjoin = array()){
        $array = array();
        $ds = new generalConsultBD();
        $datosBD = new datosBD();
        $result = $ds->GetAllData($tabla,$activo, $query, $order, $leftjoin);
        $array = $datosBD->arrDatosObj($result);
        return $array;
    }
    public function Actualizar($tabla,$campo, $valor, $nombreID ,$id){
        $param[0] = $tabla;
        $param[1] = $campo;
        $param[2] = $valor;
        $param[3] = $nombreID;
        $param[4] = $id;

        $objDB = new generalConsultBD();
        $resAct = $objDB->updateDB($param);
        return $resAct;
    }

    public function Eliminar($tabla,$nombreID,$id){
        $objDB = new generalConsultBD();
        $param[0] = $tabla;
        $param[1] = $nombreID;
        $param[2] = $id;
        return $objDB->deleteDB($param);
    }


}
