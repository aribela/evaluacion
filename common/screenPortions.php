<?php
/*
 *  © 2012 Framelova. All rights reserved. Privacy Policy
 *  Creado: 11/01/2017
 *  Por: JCRG
 *  Description: Contain necessary methods to call portions screen used as part
 *  of the template
 */

//logo page
function getLogo($level=false){
    $level=($level==true)?'../':'';
    //$logo = '<div class="logo"><a href="index.php"><img src="'.$level.'images/aguilar-29.png" /></a></div>';
    $logo = "";
    return $logo;
}

//Menu dependiendo de los permisos
function getAdminMenu(){
    $menu = '';

    $menu .= '<li id="li_inicio"><a href="index.php"><div id="divicon_sesion"><i class="icon material-icons">home</i></div><p>Inicio</p></a></li>';
    if($_SESSION['idRol'] == 1 || $_SESSION['idRol'] == 2 || $_SESSION['idRol'] == 4){
      $menu .= '<li id="li_expedientes">'
      . '<a href="formvacio.php">'
      . '<div id="divicon_sesion"><i class="icon material-icons">assignment</i></div>'
      . '<p>Form Ejemplo</p>'
      . '</a>'
      . '</li>';
      // $menu .= '<li id="li_gastos">'
      // . '<a href="gastos.php">'
      // . '<div id="divicon_sesion"><i class="icon material-icons">assignment</i></div>'
      // . '<p>Gastos</p>'
      // . '</a>'
      // . '</li>';
    }
    if($_SESSION['idRol'] == 1 || $_SESSION['idRol'] == 2){
        $menu .= '<li id="li_catalogos">'
      . '<a href="catalogos.php">'
      . '<div id="divicon_sesion"><i class="icon material-icons">assignment</i></div>'
      . '<p>Cat&aacute;logos</p>'
      . '</a>'
      . '</li>';
    }
    

    $menu .= '<li><a href="../admin/logout.php"><div id="divicon_sesion"><i class="icon material-icons-outlined">logout</i></div><p>Cerrar Sesi&oacute;n</p></a></li>';

    return $menu;
}

//footer Page
function getPiePag($level=false){
    $level=($level==true)?'../':'';
    $html = "";
    $html .= '<div class="footer_site text-muted text-center">
                   Proyecto base 2021 Todos los derechos reservados.<br>
                   Prowered by: <a href="http://framelova.com" target="_blank"><img src="'.$level.'images/framelova.png"></a>
               </div>';

    return $html;
}

//usuario
function getUsrForHeader($usrName){
  $hdr = '<div class="user text-right cursorPointer">
            <div class="dropdown">
              <div class="dropdown-toggle" data-toggle="dropdown" href="#"><span class="glyphicon glyphicon-user" aria-hidden="true"></span><strong>Usuario <span class="caret"></span>:</strong> '.$usrName.' </div>
              <ul class="dropdown-menu">
                <li><a class="user" href="../admin/logout.php">Cerrar sesi&oacute;n</a></li>
              </ul>
            </div>
          </div>
        ';
  return $hdr;
}

//header Page
function getHeaderMain($myusername, $bool){
	$html = "";
		$html .= '<header>
            		 <div class="container">
            			<div class="row">
            				<div class="col-md-9 col-sm-3 col-xs-3"> <div class="logo"><a href="index.php"><img src="../images/logo.jpg" /></a></div></div>
            				<div class="colmenu col-md-3 col-sm-3 col-xs-6">'.getUsrForHeader($myusername).'</div>
            			</div>
            	   </div>
            	</header>
            	<div class="panel-heading">
              		<div class="container">
            			<div class="row">
            				<div class="colmenu col-md-3 col-sm-3 col-xs-3">'.getLogo($bool).'</div>
            			</div>
                   </div>
             </div>';
     $html .= '<div id="fancyAyudaWeb" style="display:none;width:550px;height:300px;">
                           <div class="col-md-12 col-sm-12 col-xs-12" ><h3 id="tituloAyuda"></h3></div>
                           <div class="col-md-12 col-sm-12 col-xs-12" id="contenidoAyudaM" style="height:500px;"></div>
                </div>
                    <div id="fancyTablaContenido" style="display:none;width:550px;height:300px;">
                        <div class="col-md-12 col-sm-12 col-xs-12" ><h3 id="tituloContenido"></h3></div>
                        <div class="col-md-12 col-sm-12 col-xs-12" id="contenidoTabla" style="height:500px;"></div>
             </div>
             <div id="fancyElimCat" style="display:none;width:550px;height:100px;">
                           <div class="col-md-12 col-sm-12 col-xs-12" >
                                <h3 id="tituloElim"></h3>
                            </div>
                           <div class="col-md-12 col-sm-12 col-xs-12">
                               <input type="hidden" name="elimRegId" id="elimRegId" value="">
                                <input type="hidden" name="elimTipo" id="elimTipo" value="">
                                <div class="row">
                                    <div class="alert alert-warning" id="warningNoElim">
                                      <strong>Atenci&oacute;n!</strong> El registro no se puede eliminar porque es utilizado.
                                    </div>
                                </div>
                                <div class="row" id="contenidoElim">

                                </div>

                                <div class="row">
                                    <div class="col-md-5"></div>
                                    <div class="col-md-3">
                                        <a class="btn btn-primary" onclick="parent.$.fancybox.close();">Cerrar</a>
                                    </div>
                                    <div class="col-md-3">
                                        <a class="btn btn-primary" id="btnElimReg" onclick="eliminarRegCatalogo()">Eliminar</a>
                                    </div>
                                </div>
                           </div>
                        </div>
                ';
	return $html;
}

function getNav($menu){
  $html = "";
  $html .= '
      <nav class="navbar navbar-default" role="navigation">
          <div class="cont_menu">
              <div class="navbar-header">
                  <button type="button" class="navbar-toggle" data-toggle="collapse" data-target=".navbar-ex1-collapse">
                      <span class="sr-only">Toggle navigation</span>
                      <span class="icon-bar"></span>
                      <span class="icon-bar"></span>
                      <span class="icon-bar"></span>
                  </button>
                  <a class="navbar-brand" href="#">Men&uacute;</a>
              </div>
              <div class="collapse navbar-collapse navbar-ex1-collapse">
                  <ul class="nav navbar-nav">';
                      $html .= $menu;
                  $html .= '</ul>
              </div>
          </div>
      </nav>';
    echo  $html;
}

function getCSSRot(){
  //Obtener los colores de base de datos (del template seleccionado)
  $css = '
    :root {
      --header-bg-color: #f3912c;
      --footer-bg-color: #f3912c;
      --usertext-bg-color:#434f82;
      --navbar-bg-color: #434f82;
      --menubg-bg-color: #ccc;
      --h1-color: #434f82;
      --conticon-bg-color: #434f82;
      --conticon-border-color: #98a3d0;
      --conticon-bg-color-hover: #282f4b;
      --navbarli-hover: #434f82;
      --btn-bg-color: #98a3d0;
      --header-kool-color: #434f82;
      --btn-kool-color: #98a3d0;
    }
  ';

  // $css = '
  //   :root {
  //     --header-bg-color: #030e2d;
  //     --footer-bg-color: #030e2d;
  //   }
  // ';

  return $css;
}


//Metodo para dar enlazar todos los estilos necesarios para la pagina
function estilosPagina($level=false){
  $level=($level==true)?'../':'';
  $upd = time();

  $link = '';
  $link .= ' <link rel="icon" href="'.$level.'favicon.ico" type="image/x-icon"/>';
  $link .= ' <style type="text/css">'.getCSSRot().'</style>';
  $link .= '<link href="'.$level.'css/bootstrap.min.css" rel="stylesheet" type="text/css"/>    ';
  $link .= '<link href="'.$level.'css/style.css?upd='.$upd.'" rel="stylesheet" type="text/css" />';
  $link .= '<link href="'.$level.'css/style-responsive.css?upd='.$upd.'" rel="stylesheet" type="text/css" />';
  $link .= '<link href="'.$level.'js/fancybox/jquery.fancybox.css" rel="stylesheet" type="text/css" />';
  $link .= '<link href="'.$level.'css/alertify.min.css" rel="stylesheet" type="text/css" />';
  $link .= '<link href="'.$level.'js/boostrap-select/bootstrap-select.min.css" rel="stylesheet" type="text/css" />';
  $link .= '<link href="'.$level.'css/themes/semantic.min.css" rel="stylesheet" type="text/css"/>';
  $link .= ' <link href="'.$level.'libs/responsive_multipurpose_tabs_accordion/css/style.css" rel="stylesheet" type="text/css" />';
  $link .= ' <link href="'.$level.'css/bootstrap-duallistbox.css" rel="stylesheet" type="text/css" />';

  $link .= '
  <link href="https://fonts.googleapis.com/css2?family=Material+Icons" rel="stylesheet">
  <link href="https://fonts.googleapis.com/css2?family=Material+Icons+Outlined" rel="stylesheet">
  <link href="https://fonts.googleapis.com/css2?family=Material+Icons+Round" rel="stylesheet">
  <link href="https://fonts.googleapis.com/css2?family=Material+Icons+Sharp" rel="stylesheet">
  <link href="https://fonts.googleapis.com/css2?family=Material+Icons+Two+Tone" rel="stylesheet">
  ';

  // $link .= ' <link href="'.$level.'libs/multiupload/css/multiupload.css" rel="stylesheet" type="text/css" />';

  /*<!-- <link href="https://fonts.googleapis.com/css?family=Lato" rel="stylesheet"> -->
  <!-- <link href="../css/alertify.default.css" rel="stylesheet" type="text/css" /> -->
  <link rel="stylesheet" href="../js/bootstrap-datetimepicker/css/bootstrap-datetimepicker.min.css" />*/

  /*
  $link .= ' <link href="'.$level.'css/jquery-ui.css" rel="stylesheet" type="text/css" />';
  $link .= ' <link href="'.$level.'css/jquery.signature.css" rel="stylesheet" type="text/css" />';
  $link .= ' <link href="'.$level.'css/jquery.timepicker.min.css" rel="stylesheet" type="text/css" />';
  $link .= ' <link href="'.$level.'css/jquery.timepicker.css" rel="stylesheet" type="text/css" />';
  */
  return $link;
}

//Metodo para dar enlazar todos los scripts necesarios para la pagina
function scriptsPagina($level=false){
  $level=($level==true)?'../':'';
  $upd = time();
  $link = '';

  $link .= '<script type="text/javascript" src="'.$level.'js/jquery-1.10.2.min.js"></script>';
  $link .= '<script type="text/javascript" src="'.$level.'js/bootstrap.min.js"></script>    ';
  $link .= '<script type="text/javascript" src="'.$level.'js/jquery-ui.js"></script>';
  $link .= '<script type="text/javascript" src="'.$level.'js/fancybox/jquery.fancybox.pack.js"></script>';
  $link .= '<script type="text/javascript" src="'.$level.'js/jquery.validate.js?upd='.$upd.'"></script>';
  $link .= '<script type="text/javascript" src="'.$level.'js/additional-methods.js?upd='.$upd.'"></script>';
  $link .= '<script type="text/javascript" src="'.$level.'js/accounting.min.js"></script>';
  $link .= '<script type="text/javascript" src="'.$level.'js/alertify.js"></script>';
  $link .= '<script type="text/javascript" src="'.$level.'js/jquery_base64/jquery.base64.js"></script>';
  $link .= '<script type="text/javascript" src="'.$level.'js/boostrap-select/bootstrap-select.min.js"></script>';
  $link .= '<script>var tieneAlertify = true;</script>';


  //$link .= ' <script type="text/javascript" src="'.$level.'js/accounting.min.js?upd='.$upd.'"></script>';
  $link .= ' <script type="text/javascript" src="'.$level.'libs/responsive_multipurpose_tabs_accordion/js/jquery.multipurpose_tabcontent.js?upd='.$upd.'"></script>';

  /*<script type="text/javascript" src="../js/bootstrap-datetimepicker/js/moment-with-locales.min.js"></script>
  <script type="text/javascript" src="../js/bootstrap-datetimepicker/js/bootstrap-datetimepicker.min.js"></script>*/

  $link .= ' <script type="text/javascript" src="'.$level.'js/jquery.dataTables.min.js?upd='.$upd.'"></script>';
  $link .= ' <script type="text/javascript" src="'.$level.'js/dataTables.bootstrap.min.js?upd='.$upd.'"></script>';
  $link .= ' <script type="text/javascript" src="'.$level.'js/dataTables.fixedColumns.min.js?upd='.$upd.'"></script>';
  $link .= ' <script type="text/javascript" src="'.$level.'js/jquery.bootstrap-duallistbox.js?upd='.$upd.'"></script>'; // duallistbox
  $link .= ' <script type="text/javascript" src="'.$level.'js/jquery.serializejson.min.js?upd='.$upd.'"></script>';
  // $link .= ' <script type="text/javascript" src="'.$level.'js/jquery_base64/jquery.base64.min.js?upd='.$upd.'" type="text/javascript"></script>';
  $link .= ' <script type="text/javascript" src="'.$level.'js/jquery_base64/jquery.base64.js?upd='.$upd.'"></script>';
  $link .= ' <script type="text/javascript" src="'.$level.'js/tinymce/tinymce.min.js?upd='.$upd.'"></script>';
  // $link .= ' <script type="text/javascript" src="'.$level.'libs/multiupload/js/multiupload.js?upd='.$upd.'"></script>';
  /*
  $link .= ' <script type="text/javascript" src="'.$level.'js/jquery.timepicker.min.js?upd='.$upd.'"></script>';
  $link .= ' <script type="text/javascript" src="'.$level.'js/jquery-ui-timepicker.js?upd='.$upd.'"></script>';
  $link .= ' <script type="text/javascript" src="'.$level.'js/spanish_datapicker.js?upd='.$upd.'"></script>';
  */
  $link .= ' <script type="text/javascript" src="'.$level.'js/functionsGlobals.js?upd='.$upd.'"></script>'; //Solo para las funciones globales para todo el sitio
  $link .= ' <script type="text/javascript" src="'.$level.'js/functions.js?upd='.$upd.'"></script>';
  return $link;
}


function getFancyConcepto(){
  $html = '';

  $html .= '
  <div id="fancyConcepto" style="display:none;width:750px;height:300px;">
    <div class="col-xs-12" >
      <h3>Conceptos <span id="spanCondomino"></span></h3>
    </div>
    
    <div class="col-xs-12" id="contenidoConceptos"></div>
  </div>
  ';

  return $html;
}

?>
