<?php
$dirname = dirname(__DIR__);
include_once  $dirname.'/common/DataServices.php';

class generalConsultBD {

    //Otener todos los usuarios
    public function GetAllData($tabla,$activo, $queryExt, $order = "", $leftjoin = array()){
        $ds = new DataServices();
        $query = array();
        $param[0] = "";
        $param[1] = "";
        $where = "";
        $left = "";

        if ($activo !== "") {
          $query[] = " activo=".$activo;
        }

        $query = array_merge($query, $queryExt);

       //En caso de llevar filtro
      if(count($query) > 0){
        $wordWhere = " WHERE ";
        $setWhere = implode(" AND ", $query);
        // echo $setWhere;

        $where = $wordWhere.$setWhere;
      }

      if(count($leftjoin) > 0){
        
        $setLeft = " LEFT JOIN ".implode(" LEFT JOIN ", $leftjoin);
       
        $left = $setLeft;
      }

      // echo "tabla: ".$tabla." <br>";
        $param[0]=$tabla." ".$left." ".$where;

        if($order != ""){
            $param[1] = " ORDER BY ".$order;
        }

        // echo "<pre>"; print_r($param); echo "</pre>";

        $result = $ds->Execute("GetAllData", $param);
        $ds->CloseConnection();
        return $result;
    }

    public function updateDB($param){
        $ds = new DataServices();
        $result = $ds->Execute("updateDB", $param, false, true);
        $ds->CloseConnection();
        return $result;
    }
    // //Recuperar contrasenia por su email

    public function deleteDB($param){
      $ds = new DataServices();
      $result = $ds->Execute("deleteDB", $param, false, true);
      $ds->CloseConnection();
      return $result;
    }


    // //Obtener ids de usuarios X numero de contrato
    // public function ObtIdsUsuarioXNoContratoDB($noContratos)
    // {
    //   $ds = new DataServices();
    //   $param[0]= $noContratos;

    //   $result = $ds->Execute("ObtIdsUsuarioXNoContratoDB", $param);
    //   $ds->CloseConnection();
    //   return $result;
    // }

}
