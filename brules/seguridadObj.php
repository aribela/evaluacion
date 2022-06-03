<?php

class seguridadObj {
    
    //Encriptar
    public function encriptarCadena($cadena){
         $semilla = "__!?/?!__".rand(100,1000000);
         $cad = rtrim(strtr(base64_encode($cadena.$semilla), '+/', '-_'), '=');
         
         return $cad;
    }

    //Desencriptar
    public function desencriptarCadena($cadena){
        $cad = base64_decode(str_pad(strtr($cadena, '-_', '+/'), strlen($cadena) % 4, '=', STR_PAD_RIGHT));        
        $arrCad = explode("__!?/?!__", $cad); //Remover semilla
        $cad = $arrCad[0];

        return $cad;
    }    
    //Metodo para obtener url completa despues del simbolo ?
    public function obtUrlCompleta(){
        $actual_url = (isset($_SERVER['HTTPS']) ? "https" : "http") . "://$_SERVER[HTTP_HOST]$_SERVER[REQUEST_URI]";
        $arrUrl = explode("?", $actual_url);
        return $arrUrl;
    }
    //Metodo para separa url por & y setearlos en get
    public function separarUrlObtGet($cadena){
      $arrCadena = explode("&", $cadena);

      if(count($arrCadena)>0){
        foreach ($arrCadena as $elem){            
            $elemExp = explode("=", $elem);            
            //Crear get ($elemExp[0] = key, $elemExp[1]=valor)
            $_GET[$elemExp[0]] = $elemExp[1];        
        }
      }
    }
    
}

?>