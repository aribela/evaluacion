<?php
$dirname = dirname(__DIR__);

class configuracionesGridObj {

    //Define el grid
    protected function defineGrid($grid, $ds, $activar = 0)
    {
        // $grid->id // Es el nombre del grid
        // echo "<pre>";
        // print_r($grid->id);
        // echo "</pre>";
        //create and define grid
        $grid->scriptFolder = "../brules/KoolControls/KoolGrid";
        $grid->styleFolder="office2010blue";
        $grid->Width = "760px";
        $grid->AllowScrolling = false;
        //$grid->MasterTable->Height = "540px";
        //$grid->MasterTable->ColumnWidth = "130px";
        $grid->RowAlternative = true;
        $grid->AjaxEnabled = true;
        $grid->AjaxLoadingImage =  "../brules/KoolControls/KoolAjax/loading/5.gif";
        $grid->Localization->Load("../brules/KoolControls/KoolGrid/localization/es.xml");
        $grid->AllowInserting = true;
        $grid->AllowEditing = true;
        $grid->AllowDeleting = true;
        $grid->AllowSorting = true;
        $grid->ColumnWrap = true;
        $grid->AllowResizing = true;
        $grid->MasterTable->DataSource = $ds;
        $grid->MasterTable->AutoGenerateColumns = false;
        $grid->MasterTable->Pager = new GridPrevNextAndNumericPager();
        $grid->MasterTable->Pager->ShowPageSize = true;
        $grid->MasterTable->Pager->PageSizeOptions = "10,25,50,100,150";

        if($_SESSION['idRol']==1 || $_SESSION['idRol']==2 || $_SESSION['idRol']==5 || $_SESSION['idRol']==4){
            //Show Function Panel
            if ($grid->id == "usuariosGrid") {
                $grid->MasterTable->ShowFunctionPanel = true;
                $grid->MasterTable->FunctionPanel->ShowInsertButton = false;
            }elseif ($grid->id == "rolGrid") {
                $grid->MasterTable->ShowFunctionPanel = true;
                $grid->MasterTable->FunctionPanel->ShowInsertButton = false;
            }elseif ($grid->id == "comunicados") {
                $grid->MasterTable->ShowFunctionPanel = true;
                $grid->MasterTable->FunctionPanel->ShowInsertButton = false;
            }elseif ($grid->id == "casos") {
                $grid->MasterTable->ShowFunctionPanel = true;
                $grid->MasterTable->FunctionPanel->ShowInsertButton = false;
                $grid->MasterTable->Pager->PageSize = 50;
                $grid->Width = "900px";
                $grid->MasterTable->Height = "1500px";
                $grid->MasterTable->Pager->Position = "top+bottom";
	            $grid->AllowScrolling = true;
            }elseif ($grid->id == "caso_acciones") {
                $grid->MasterTable->ShowFunctionPanel = true;
                $grid->MasterTable->FunctionPanel->ShowInsertButton = false;
                $grid->ClientSettings->ClientEvents["OnPageChange"] = "Handle_OnPageChange";
                $grid->ClientSettings->ClientEvents["OnLoad"] = "Handle_OnLoad";
                $grid->MasterTable->Height = "900px";
	            $grid->AllowScrolling = true;
            }elseif ($grid->id == "accion_gastos") {
                $grid->MasterTable->ShowFunctionPanel = true;
                $grid->MasterTable->FunctionPanel->ShowInsertButton = false;
            }elseif ($grid->id == "gastosGrid") {
                $grid->MasterTable->ShowFunctionPanel = true;
                $grid->MasterTable->FunctionPanel->ShowInsertButton = false;
                $grid->MasterTable->Pager->PageSize = 25;
            }elseif ($grid->id == "tareas") {
                $grid->MasterTable->ShowFunctionPanel = true;
                $grid->MasterTable->FunctionPanel->ShowInsertButton = false;
                $grid->MasterTable->Height = "1000px";
	            $grid->AllowScrolling = true;
	            $grid->MasterTable->Pager->Position = "top+bottom";
            }
            elseif ($grid->id == "digitales") {
                $grid->MasterTable->ShowFunctionPanel = true;
                $grid->MasterTable->FunctionPanel->ShowInsertButton = false;
                $grid->ClientSettings->ClientEvents["OnPageChange"] = "Handle_OnPageChange";
                $grid->ClientSettings->ClientEvents["OnLoad"] = "Handle_OnLoad";
                $grid->ClientSettings->ClientEvents["OnInit"] = "Handle_OnInit";
            }
            elseif ($grid->id == "pagos" || $grid->id == "adicionales") {
                $grid->MasterTable->ShowFunctionPanel = true;
                $grid->MasterTable->FunctionPanel->ShowInsertButton = false;
            }
            elseif ($grid->id == "cuentas") {
                $grid->MasterTable->ShowFunctionPanel = true;
                $grid->MasterTable->FunctionPanel->ShowInsertButton = false;
            }
            /*elseif ($grid->id == "grid_actividades") {
                $grid->MasterTable->ShowFunctionPanel = true;
                $grid->MasterTable->FunctionPanel->ShowInsertButton = false;
                $grid->AllowScrolling = true;
                // $grid->MasterTable->Pager->PageSize = 25;
            }*/
            
            else {
                $grid->MasterTable->ShowFunctionPanel = true;
            }
        }
        if ($grid->id == "grid_actividades") {
            $grid->MasterTable->ShowFunctionPanel = true;
            $grid->MasterTable->FunctionPanel->ShowInsertButton = false;
            $grid->AllowScrolling = true;
            $grid->ClientSettings->ClientEvents["OnInit"] = "Handle_OnInit";
            $grid->ClientSettings->ClientEvents["OnLoad"] = "Handle_OnLoad";
            $grid->MasterTable->Height = "600px";
            // $grid->MasterTable->Pager->PageSize = 25;
        }
        if ($grid->id == "grid_tareas") {
            $grid->MasterTable->ShowFunctionPanel = true;
            $grid->MasterTable->FunctionPanel->ShowInsertButton = false;
            $grid->AllowScrolling = true;
            $grid->ClientSettings->ClientEvents["OnInit"] = "Handle_OnInit";
            $grid->ClientSettings->ClientEvents["OnLoad"] = "Handle_OnLoad";
            $grid->MasterTable->Height = "400px";
            // $grid->MasterTable->Pager->PageSize = 25;
        }
        //Insert Settings
        $grid->MasterTable->InsertSettings->Mode = "Form";
        $grid->MasterTable->EditSettings->Mode = "Form";
        $grid->MasterTable->InsertSettings->ColumnNumber = 1;
        $grid->ClientSettings->ClientEvents["OnRowConfirmEdit"] = "Handle_OnRowConfirmEdit";
        $grid->ClientSettings->ClientEvents["OnConfirmInsert"] = "Handle_OnConfirmInsert";
        $grid->ClientSettings->ClientEvents["OnRowDelete"] = "Handle_OnRowDelete";

        if($activar > 0){
            // $grid->AllowMultiSelecting = true;// Allow multi row selecting
	        $grid->KeepSelectedRecords = true;//Keep selected records cross page.
        }

        switch ($grid->id) {
            case 'cat_cursos':
            case 'comunicados':
                $grid->Width = "960px";
            break;
            case 'gastosGrid':
                $grid->Width = "960px";
            break;
            /*case 'patentes_competencia':
                $grid->Width = "960px";
            break;
            case 'uso_estadisticas':
                $grid->Width = "1000px";
                $grid->MasterTable->ShowFunctionPanel = false; //Show Function Panel
                // $grid->MasterTable->Pager->Position = "top";//Show both pager
                // $grid->MasterTable->FunctionPanel->ShowInsertButton = false;
                // $grid->MasterTable->FunctionPanel->ShowRefreshButton = true;
            break;
            // case 'cat_tasas':
            //     if($_SESSION['idRol']==1 || $_SESSION['idRol']==2){
            //         $grid->MasterTable->ShowFunctionPanel = true; //Show Function Panel
            //     }
            // break;*/
        }
    }

    //define la columna del grid
    protected function defineColumn($grid, $name_field, $name_header, $visible=true, $read_only=false, $validator=0, $field_type="", $width = "90px", $filtroTitular = "",$pantalla=0){
        $column = new gridboundcolumn();

        if($grid->id=="usuariosGrid"){
            if($name_field == 'idRol') {
                $column = new GridDropDownColumn();
                $rolObj = new rolesObj();
                $rolArr = $rolObj->GetAllRoles();
                $column->AddItem('-- Seleccionar --',NULL);
                foreach($rolArr as $rolTmp){
                    $column->AddItem($rolTmp->rol,$rolTmp->idRol);
                }
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal");//Only show 3 chosen
            }
            elseif($name_field == 'activo') {
                $column = new GridBooleanColumn();
                $column->UseCheckBox = true;
                // $column->CheckBox = true;
            }elseif ($name_field == 'nombre' || $name_field == 'email') {
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal","Contain");//Only show 3 chosen
            }
            else{
                $column = new gridboundcolumn();
            }
        }
        elseif($grid->id=="clientes"){
            if ($name_field == 'nombre') {
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Contain");//Only show 3 chosen
            }
            elseif ($name_field == 'empresa') {
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Contain");//Only show 3 chosen
            }
            elseif ($name_field == 'direccion') {
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Contain");//Only show 3 chosen
            }
        }
        elseif($grid->id=="digitales"){
            if ($name_field == 'tipo') {
                $column = new GridDropDownColumn();
                $column->AddItem('Escrito',1);
                $column->AddItem('Expediente',2);
                $column->AddItem('Audiencias',3);
                $column->AddItem('Otros',4);
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal");//Only show 3 chosen
            }
            elseif ($name_field == 'nombre') {
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal","Contain");//Only show 3 chosen
            }
        }
        elseif($grid->id=="cat_comunicados"){
            if ($name_field == 'activo') {
                $column = new GridBooleanColumn();
                $column->UseCheckBox = true;  
            }
            elseif ($name_field == 'descripcionCorta') {
                $column = new GridTextAreaColumn();
                $column->BoxHeight = "50px";
            }
            elseif ($name_field == 'contenido') {
                $column = new GridTextAreaColumn();
                $column->BoxHeight = "100px";
                $column->AllowHtmlRender = true;
            }
        }
        elseif($grid->id=="cat_tipo_casos"){
            if ($name_field == 'activo') {
                $column = new GridDropDownColumn();
                $column->AddItem('No',0);
                $column->AddItem('Si',1);
            }
        }
        elseif($grid->id=="cat_conceptos"){
            if ($name_field == 'activo') {
                $column = new GridDropDownColumn();
                $column->AddItem('No',0);
                $column->AddItem('Si',1);
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal");//Only show 3 chosen
            }elseif($name_field == 'tipo'){
                $column = new GridDropDownColumn();
                $column->AddItem('Entradas',1);
                $column->AddItem('Salidas',2);
                $column->AddItem('Gastos admin',3);
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal");//Only show 3 chosen
            }elseif($name_field == 'nombre'){
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal","Contain");//Only show 3 chosen
            }
        }elseif ($grid->id == "gastosGrid") {
            if($name_field == 'idRol') {
                $column = new GridDropDownColumn();
                $rolObj = new rolesObj();
                $rolArr = $rolObj->GetAllRoles();
                $column->AddItem('-- Seleccionar --',NULL);
                foreach($rolArr as $rolTmp){
                    $column->AddItem($rolTmp->rol,$rolTmp->idRol);
                }
                // $column->AllowFiltering = true;
                // $column->FilterOptions  = array("No_Filter","Equal");//Only show 3 chosen
            }elseif ($name_field == 'saldo') {
                $column = new GridCustomColumn();
                $column->ItemTemplate = '
                $ {saldo}';
                $column->Align = "right";
            }
            elseif ($name_field == 'saldoPeriodo') {
                $column = new GridCustomColumn();
                $column->ItemTemplate = '
                $ {saldoPeriodo}';
                $column->Align = "right";
            }elseif ($name_field == 'totalCargos') {
                $column = new GridCustomColumn();
                $column->ItemTemplate = '
                $ -{totalCargos}';
                $column->Align = "right";
            }elseif ($name_field == 'totalAbonos') {
                $column = new GridCustomColumn();
                $column->ItemTemplate = '
                $ {totalAbonos}';
                $column->Align = "right";
            }elseif ($name_field == 'nombre') {
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal","Contain");//Only show 3 chosen
            }
            elseif ($name_field == 'numAbogado') {
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal","Contain");//Only show 3 chosen
            }
        }elseif ($grid->id == "casos") {
            if ($name_field == 'titular') {//Nombre del abogado responsable
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal","Contain");//Only show 3 chosen
            }
            elseif ($name_field == 'estatusId') {
                $column = new GridDropDownColumn();
                
                $column->AddItem('-- Seleccionar --',NULL);
                
                if($pantalla === 0){//Expedientes
                    $column->AddItem('Activo',1);
                    $column->AddItem('Prospecto',5);    
                }
                elseif($pantalla === 1){//historico expedientes
                    $column->AddItem('Suspendido',2);
                    $column->AddItem('Baja',3);
                    $column->AddItem('Terminado',4);
                }

                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal");//Only show 3 chosen
            }
            elseif ($name_field == 'saludExpediente') {
                $column = new GridCustomColumn();
                $column->ItemTemplate = '<div class="{clasesalud}">{saludExpediente}</div>';
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal","Contain");//Only show 3 chosen
            }
            elseif ($name_field == 'titular2') {//Nombre del titular rol superadmin
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal","Contain");//Only show 3 chosen
                if($filtroTitular != ""){
                    //JGP 08/01/22 comente porque estaba causando que ya no servia ningun filtro
                    //$column->Filter = array("Value"=>$filtroTitular,"Exp"=>"Contain"); // Make col filter those row has field value greater than 123
                }
            }
            elseif ($name_field == 'cliente') {
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal","Contain");//Only show 3 chosen
            }
            elseif ($name_field == 'representado') {
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal","Contain");//Only show 3 chosen
            }
            elseif ($name_field == 'contrario') {
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal","Contain");//Only show 3 chosen
            }
            elseif ($name_field == 'cobro') {
                $column = new GridDropDownColumn();

                $column->AddItem('-- Seleccionar --',NULL);
                $column->AddItem("Probono",1);
                $column->AddItem("Sin pago temporal",2);
                $column->AddItem("Pago activo",3);
                $column->AddItem("Liquidado",4);
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal");//Only show 3 chosen
            }
            elseif ($name_field == 'modoCobro') {
                $column = new GridDropDownColumn();

                $column->AddItem('-- Seleccionar --',NULL);
                $column->AddItem("Monto fijo pago unico",1);
                $column->AddItem("Pago inicial e igualas mensuales",2);
                $column->AddItem("Pago inicial y pago sobre avances",3);
                $column->AddItem("Monto fijo en pagos",4);
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal");//Only show 3 chosen
            }
            elseif ($name_field == 'aka') {
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal","Contain");//Only show 3 chosen
            }
            elseif ($name_field == 'tipocaso') {
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal","Contain");//Only show 3 chosen
            }
            elseif ($name_field == 'numAbogado') {
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal","Contain");//Only show 3 chosen
            }
            elseif ($name_field == 'idCaso') {
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal","Contain");//Only show 3 chosen
            }
	    elseif ($name_field == 'numExpediente') {
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal","Contain");//Only show 3 chosen
            }
            elseif ($name_field == 'numExpedienteJuzgado') {
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal","Contain");//Only show 3 chosen
            }
            elseif($name_field == 'parteId') {
                $column = new GridDropDownColumn();
                $rolObj = new catPartesObj();
                $rolArr = $rolObj->ObtPartes("", "nombre");
                $column->AddItem('-- Seleccionar --',NULL);
                foreach($rolArr as $rolTmp){
                    $column->AddItem($rolTmp->nombre,$rolTmp->idParte);
                }
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal");//Only show 3 chosen
            }
            elseif($name_field == 'materiaId') {
                $column = new GridDropDownColumn();
                $rolObj = new catMateriasObj();
                $rolArr = $rolObj->ObtMaterias("", "nombre");
                $column->AddItem('-- Seleccionar --',NULL);
                foreach($rolArr as $rolTmp){
                    $column->AddItem($rolTmp->nombre,$rolTmp->idMateria);
                }
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal");//Only show 3 chosen
            }
            elseif($name_field == 'juicioId') {
                $column = new GridDropDownColumn();
                $rolObj = new catJuiciosObj();
                $rolArr = $rolObj->ObtJuicios("", -1, "nombre");
                $column->AddItem('-- Seleccionar --',NULL);
                foreach($rolArr as $rolTmp){
                    $column->AddItem($rolTmp->nombre,$rolTmp->idJuicio);
                }
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal");//Only show 3 chosen
            }
            elseif($name_field == 'juzgadoId') {
                $column = new GridDropDownColumn();
                $rolObj = new catJuzgadosObj();
                $rolArr = $rolObj->ObtJuzgados("", -1, "nombre");
                $column->AddItem('-- Seleccionar --',NULL);
                foreach($rolArr as $rolTmp){
                    $column->AddItem($rolTmp->nombre,$rolTmp->idJuzgado);
                }
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal");//Only show 3 chosen
            }
            elseif($name_field == 'seleccionar') {
                $column = new GridCustomColumn();
                $column->ItemTemplate = '<input type="checkbox" name="seleccionar_{idCaso}" id="seleccionar_{idCaso}" value="{idCaso}" onchange="cambiaSeleccionaCaso({idCaso})">';

            }
        }elseif($grid->id == "caso_acciones"){
            if ($name_field == 'comentarios') {
                $column = new GridCustomColumn();
                $column->ItemTemplate = '<div class="wrapper-col-contenido">
                        <div class="small-col-contenido">{comentarios}</div>
                        <a class="linktoogle-contenido" href="#">+</a>
                        </div>';

                //$column = new GridBoundColumn();
                //$col->AllowHtmlRender = true;
                //$column->MaxLength = 200;
                //$grid->MasterTable->AddColumn($column);
            }
            elseif($name_field == 'estatusId') {
                $column = new GridDropDownColumn();
                
                $column->AddItem('-- Seleccionar --',NULL);
                $column->AddItem('Por realizar',1);
                $column->AddItem('En proceso',2);
                $column->AddItem('Terminado',4);
                $column->AddItem('Espero instrucciones',3);
                
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal");//Only show 3 chosen
            }
        }
        elseif($grid->id=="cat_partes"){
            if ($name_field == 'activo') {                
                $column = new GridBooleanColumn();
                $column->UseCheckBox = true;                
            }
        }
        elseif($grid->id=="cat_materias"){
            if ($name_field == 'activo') {                
                $column = new GridBooleanColumn();
                $column->UseCheckBox = true;                
            }elseif ($name_field == 'tieneAcciones') {                
                $column = new GridBooleanColumn();
                $column->UseCheckBox = true;                
            }
        }
        elseif($grid->id=="cat_juicios"){
            if ($name_field == 'activo') {                
                $column = new GridBooleanColumn();
                $column->UseCheckBox = true;                
            }
            elseif($name_field == 'materiaId') {
                $column = new GridDropDownColumn();
                $rolObj = new catMateriasObj();
                $rolArr = $rolObj->ObtMaterias("", "nombre");
                $column->AddItem('-- Seleccionar --',NULL);
                foreach($rolArr as $rolTmp){
                    $column->AddItem($rolTmp->nombre,$rolTmp->idMateria);
                }
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal");//Only show 3 chosen
            }
        }
        elseif($grid->id=="cat_distritos"){
            if ($name_field == 'activo') {                
                $column = new GridBooleanColumn();
                $column->UseCheckBox = true;                
            }
        }
        elseif($grid->id=="cat_metodos"){
            if ($name_field == 'activo') {                
                $column = new GridBooleanColumn();
                $column->UseCheckBox = true;                
            }
            elseif ($name_field == 'requiereBanco') {                
                $column = new GridBooleanColumn();
                $column->UseCheckBox = true;                
            }
        }
        elseif($grid->id=="cat_bancos"){
            if ($name_field == 'activo') {                
                $column = new GridBooleanColumn();
                $column->UseCheckBox = true;                
            }
        }
        elseif($grid->id=="cat_juzgados"){
            if ($name_field == 'activo') {                
                $column = new GridBooleanColumn();
                $column->UseCheckBox = true;                
            }
            elseif($name_field == 'distritoId') {
                $column = new GridDropDownColumn();
                $rolObj = new catDistritosObj();
                $rolArr = $rolObj->ObtDistritos();
                $column->AddItem('-- Seleccionar --',NULL);
                foreach($rolArr as $rolTmp){
                    $column->AddItem($rolTmp->nombre,$rolTmp->idDistrito);
                }
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal");//Only show 3 chosen
            }
        }
        elseif($grid->id=="cat_acciones"){
            if ($name_field == 'activo') {                
                $column = new GridBooleanColumn();
                $column->UseCheckBox = true;                
            }
            elseif($name_field == 'materiaId') {
                $column = new GridDropDownColumn();
                $rolObj = new catMateriasObj();
                $rolArr = $rolObj->ObtMaterias("", "nombre", 1);
                $column->AddItem('-- Seleccionar --',NULL);
                foreach($rolArr as $rolTmp){
                    $column->AddItem($rolTmp->nombre,$rolTmp->idMateria);
                }
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal");//Only show 3 chosen
            }
        }
        elseif($grid->id=="grid_actividades"){
            if ($name_field == 'activo') {                
                $column = new GridBooleanColumn();
                $column->UseCheckBox = true;                
            }
            elseif($name_field == 'estatusId') {
                $column = new GridDropDownColumn();
                
                $column->AddItem('-- Seleccionar --',NULL);
                $column->AddItem('Por realizar',1);
                $column->AddItem('En proceso',2);
                $column->AddItem('Terminado',4);
                $column->AddItem('Espero instrucciones',3);
                
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal");//Only show 3 chosen
            }
            elseif($name_field == 'nombreAbogado' || $name_field == 'numExpediente' || $name_field == 'nombre' || $name_field == 'nombreCliente' || $name_field == 'idCaso' || $name_field == 'idAccion' || $name_field == 'fechaRealizado2' || $name_field == 'diferenciaTerminado' || $name_field == 'numExpedienteJuzgado' || $name_field == 'NombreTitular' || $name_field == 'ajeno') {
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal","Contain");//Only show 3 chosen
            }
            elseif($name_field == 'fechaCreacion2') {
                $column = new GridCustomColumn();
                $column->ItemTemplate = '{fechaCreacion2}';
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal","Contain");//Only show 3 chosen
            }
            elseif($name_field == 'fechaCompromiso2') {
                $column = new GridCustomColumn();
                $column->ItemTemplate = '<div class="{clasevencido}">{fechaCompromiso2}</div>';
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal","Contain");//Only show 3 chosen
            }
            elseif($name_field == 'fechaCompromiso') {
                $column = new GridCustomColumn();
                $column->ItemTemplate = '{fechaCompromiso}';
                $column->Sort = 1;
                // $column->AllowFiltering = true;
                // $column->Sort = 1;
                // $column->FilterOptions  = array("No_Filter","Equal","Contain");//Only show 3 chosen
            }
            elseif($name_field == 'fechaAct2') {
                $column = new GridCustomColumn();
                $column->ItemTemplate = '{fechaAct2}';
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal","Contain");//Only show 3 chosen
            }
            elseif($name_field == 'importanciaTexto'){
                $column = new GridDropDownColumn();
                
                $column->AddItem('-- Seleccionar --',NULL);
                $column->AddItem('Normal','Normal');
                $column->AddItem('Media','Media');
                $column->AddItem('Alta','Alta');

                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal");//Only show 3 chosen
            }
            elseif($name_field == 'tipoTexto'){
                $column = new GridDropDownColumn();
                
                $column->AddItem('-- Seleccionar --',NULL);
                $column->AddItem('De fondo','De fondo');
                $column->AddItem('Seguimiento','Seguimiento');
                $column->AddItem('Audiencias','Audiencias');
                $column->AddItem('Término','Término');
                $column->AddItem('Citaciones y Emplazamientos','Citaciones y Emplazamientos');

                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal");//Only show 3 chosen
            }
        }
        elseif($grid->id == "tareas"){
            if ($name_field == 'comentarios2') {
                $column = new GridCustomColumn();
                //$column->ItemTemplate = '
                //<div class="contenido_truncarx">{comentarios2}</div>';

                $column->ItemTemplate = '<div class="wrapper-col-contenido">
                        <div class="small-col-contenido">{comentarios2}</div>
                        <a class="linktoogle-contenido" href="#">+</a>
                        </div>';
            }
            elseif($name_field == 'numAbogado' || $name_field == 'nombreResponsable'
                || $name_field == 'idTarea' || $name_field == 'fechaAlta2'){
                
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal","Contain");//Only show 3 chosen
            }
            
            elseif ($name_field == 'reporte') {
                $column = new GridCustomColumn();
                $column->ItemTemplate = '<div class="wrapper-col-contenido">
                        <div class="small-col-contenido">{reporte}</div>
                        <a class="linktoogle-contenido" href="#">+</a>
                        </div>';
            }
            elseif($name_field == 'estatusId') {
                $column = new GridDropDownColumn();
                
                $column->AddItem('-- Seleccionar --',NULL);
                $column->AddItem('Por realizar',1);
                $column->AddItem('En proceso',2);
                $column->AddItem('Terminado',4);
                $column->AddItem('Espero instrucciones',3);
                
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal");//Only show 3 chosen
            }
            elseif($name_field == 'fechaCompromiso2') {
                $column = new GridCustomColumn();
                $column->ItemTemplate = '<div class="{clasevencido}">{fechaCompromiso2}</div>';
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal","Contain");//Only show 3 chosen
            }
        }
        elseif($grid->id=="grid_tareas"){
            if ($name_field == 'activo') {                
                $column = new GridBooleanColumn();
                $column->UseCheckBox = true;                
            }
            elseif($name_field == 'estatusId') {
                $column = new GridDropDownColumn();
                
                $column->AddItem('-- Seleccionar --',NULL);
                $column->AddItem('Por realizar',1);
                $column->AddItem('En proceso',2);
                $column->AddItem('Terminado',4);
                $column->AddItem('Espero instrucciones',3);
                
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal");//Only show 3 chosen
            }
            elseif($name_field == 'nombreCreador' || $name_field == 'nombreResponsable' || $name_field == 'numExpediente' || $name_field == 'nombre' || $name_field == 'nombreCliente' || $name_field == 'idCaso' || $name_field == 'idTarea' || $name_field == 'fechaRealizado2' || $name_field == 'diferenciaTerminado' || $name_field == 'numExpedienteJuzgado' || $name_field == 'NombreTitular' || $name_field == 'ajeno') {
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal","Contain");//Only show 3 chosen
            }
            elseif($name_field == 'fechaCompromiso2') {
                $column = new GridCustomColumn();
                $column->ItemTemplate = '<div class="{clasevencido}">{fechaCompromiso2}</div>';
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal","Contain");//Only show 3 chosen
            }
            elseif($name_field == 'fechaCompromiso') {
                $column = new GridCustomColumn();
                $column->ItemTemplate = '{fechaCompromiso}';
                $column->Sort = 1;
                
            }
            elseif($name_field == 'importanciaTexto'){
                $column = new GridDropDownColumn();
                
                $column->AddItem('-- Seleccionar --',NULL);
                $column->AddItem('Normal','Normal');
                $column->AddItem('Media','Media');
                $column->AddItem('Alta','Alta');

                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal");//Only show 3 chosen
            }
            elseif($name_field == 'tipoTexto'){
                $column = new GridDropDownColumn();
                
                $column->AddItem('-- Seleccionar --',NULL);
                $column->AddItem('Administrativa','Administrativa');
                $column->AddItem('Otros','Otros');
                // $column->AddItem('Audiencias','Audiencias');
                // $column->AddItem('Término','Término');
                // $column->AddItem('Citaciones y Emplazamientos','Citaciones y Emplazamientos');

                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal");//Only show 3 chosen
            }
        }
        elseif($grid->id=="pagos" || $grid->id=="adicionales"){
            if($name_field == 'metodoId') {
                $column = new GridDropDownColumn();
                $rolObj = new catMetodosObj();
                $rolArr = $rolObj->ObtMetodos ("", "nombre");
                $column->AddItem('-- Seleccionar --',NULL);
                foreach($rolArr as $rolTmp){
                    $column->AddItem($rolTmp->nombre,$rolTmp->idMetodo);
                }
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal");//Only show 3 chosen
            }
            elseif($name_field == 'bancoId') {
                $column = new GridDropDownColumn();
                $rolObj = new catBancosObj();
                $rolArr = $rolObj->ObtBancos ("", "nombre");
                $column->AddItem('-- Seleccionar --',NULL);
                foreach($rolArr as $rolTmp){
                    $column->AddItem($rolTmp->nombre,$rolTmp->idBanco);
                }
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal");//Only show 3 chosen
            }
            elseif ($name_field == 'monto') {
                $column = new GridCustomColumn();
                $column->ItemTemplate = '
                $ {monto}';
                $column->Align = "right";
            }
        }
        elseif($grid->id=="cuentas"){
            if($name_field == 'clienteId') {
                $column = new GridDropDownColumn();
                $rolObj = new clientesObj();
                $rolArr = $rolObj->ObtClientes ("", "nombre");
                $column->AddItem('-- Seleccionar --',NULL);
                foreach($rolArr as $rolTmp){
                    $column->AddItem($rolTmp->nombre,$rolTmp->idCliente);
                }
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal");//Only show 3 chosen
            }
            elseif($name_field == 'casoId') {
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal","Contain");//Only show 3 chosen
            }
            elseif($name_field == 'clienteId') {
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal","Contain");//Only show 3 chosen
            }
            elseif($name_field == 'tipoCobro') {
                $column = new GridDropDownColumn();

                $column->AddItem('-- Seleccionar --',NULL);
                $column->AddItem("Probono",1);
                $column->AddItem("Sin pago temporal",2);
                $column->AddItem("Pago activo",3);
                $column->AddItem("Liquidado",4);
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal");//Only show 3 chosen
            }
            elseif ($name_field == 'planPagos') {
                $column = new GridDropDownColumn();

                $column->AddItem('-- Seleccionar --',NULL);
                $column->AddItem("Monto fijo pago unico",1);
                $column->AddItem("Pago inicial e igualas mensuales",2);
                $column->AddItem("Pago inicial y pago sobre avances",3);
                $column->AddItem("Monto fijo en pagos",4);
                $column->AllowFiltering = true;
                $column->FilterOptions  = array("No_Filter","Equal");//Only show 3 chosen
            }

        }
        /*if($grid->id=="patentes_competencia"){
            if($name_field == 'esPatente') {
                $column = new GridDropDownColumn();
                $column->AddItem('No',0);
                $column->AddItem('Si',1);
            }
            elseif ($name_field == 'presentacion') {
                $column = new GridTextAreaColumn();
                $column->AllowHtmlRender = false;
                $column->BoxHeight = "50px";
            }
            elseif ($name_field == 'cicloTratamiendo') {
                $column = new GridTextAreaColumn();
                $column->AllowHtmlRender = false;
                $column->BoxHeight = "50px";
            }
        }

        if($grid->id=="uso_estadisticas"){
            if($name_field == 'usuarioId') {
                $column = new GridDropDownColumn();
                $usuarioObj = new usuariosObj();
                $colUsuarios = $usuarioObj->obtTodosUsuarios();
                foreach($colUsuarios as $usuarioTmp)
                {
                    $column->AddItem($usuarioTmp->nombre, $usuarioTmp->idUsuario);
                }
            }
        }*/


        //Valida si es requerido
        if($validator > 0){
            $column->addvalidator($this->GetValidator ($validator));
        }

        //Tipo de validacion
        if($field_type != ""){
            $column->addvalidator($this->GetValidatorFieldType($field_type));
        }

        $column->Visible = $visible;
        $column->DataField = $name_field;
        $column->HeaderText = $name_header;
        $column->ReadOnly = $read_only;
        $column->Width = $width;
        $grid->MasterTable->AddColumn($column);
    }

    //validar campo
    private function GetValidator($type){
        switch ($type) {
            case 1: //required
                $validator = new RequiredFieldValidator();
                $validator->ErrorMessage = "Campo requerido";
                return $validator;
                break;
        }
    }

    //valido el tipo del campo
    private function GetValidatorFieldType($field_type){
        switch ($field_type) {
            case "INT":
                $validatorTmp = new RegularExpressionValidator();
                $validatorTmp->ValidationExpression = "/^([0-9])+$/";
                $validatorTmp->ErrorMessage = "Campo tipo entero";
                return $validatorTmp;
            break;
            case "FLOAT":
                $validatorTmp = new RegularExpressionValidator();
                $validatorTmp->ValidationExpression = "/^([.0-9])+$/";
                $validatorTmp->ErrorMessage = "Campo tipo flotante";
                return $validatorTmp;
            break;
            case "EMAIL":
                $validatorTmp = new RegularExpressionValidator();
                $validatorTmp->ValidationExpression = "/^\w+@[a-zA-Z_]+?\.[a-zA-Z]{2,3}$/";
                $validatorTmp->ErrorMessage = "Campo tipo email";
                return $validatorTmp;
            break;
        }
    }

    //define la columna de acciones
    protected function defineColumnEdit($grid)
    {
        $column = new GridCustomColumn();
        if($_SESSION['idRol']==1 || $_SESSION['idRol']==2){
            if($grid->id == "rolGrid"){
                $column->ItemTemplate = '
                <a class="grid_edit" onclick="grid_edit(this)" href="javascript:void 0" title="Editar"></a>';
            }
            elseif ($grid->id == "usuariosGrid") {
                $column->ItemTemplate = '
                <a href="frmUsuario.php?id={idUsuario}" class="grid_edit"></a>'
                . '<a class="grid_delete btnFancy" onclick="verificaUsoTabla(\'usuarios\', {idUsuario})" href="#fancyElimCat" title="Eliminar"></a>';
            }elseif($grid->id == "comunicados"){
                $column->ItemTemplate = '<a class="grid_edit" onclick="grid_edit(this)" href="javascript:void 0" title="Editar"></a>';
                //<a class="grid_edit" onclick="edicionGrid(\'comunicados\', {idComunicado});" href="javascript:void 0" title="Editar"></a>'
                //    .' <a class="grid_delete" onclick="grid_delete(this)" href="javascript:void 0" title="Eliminar"></a>'
                // . '<a class="btnDesactivarUsuario" onclick="muestraDesactivarUsuario({idUsuario},\'usuariosGrid\',\'usuarios\',{activo})"  href="#fancyDesactivarUsuario" title="Activar/Desactivar usuario"><img src="../images/{nombreImg}" class="iconoDesactivar" ></a>'
                ;
            }
            elseif($grid->id == "casos"){
                $link = '';
                $link .= '<a class="grid_edit" onclick="edicionGrid(\'casos\', {idCaso});" href="javascript:void 0" title="Editar"></a>
                <a class="grid_price" target="_blank" href="cuentasxcobrar.php?expedienteId={idCaso}" title="Cuentas por cobrar"></a>
                ';
                // $link .= '&nbsp;&nbsp;<a class="" onclick="sincronizarEventos({clienteId}, {idCaso}, \'{cliente}\');" href="javascript:void 0" title="Sincronizar"><img src="../images/iconos/iconos_grid/actualizar.png"></a>';
                $column->ItemTemplate = $link;
            }
            elseif($grid->id == "caso_acciones"){
                $link = '';
                // $link .= '<a href="#" data-toggle="modal" data-target="#popup_modalCrearAccion" class="agregarAccion" title="Editar acci&oacute;n" idAccion="{idAccion}"><img width="16px" src="../images/iconos/iconos_grid/editar.png"></a>';
                $link .= '<a class="grid_edit" target="_blank" href="actividad.php?expId={casoId}&actId={idAccion}" title="Ver detalle acci&oacute;n"></a>';
                // $link .= '&nbsp; <a href="javascript:void(0);" onclick="popupCreaEditaGasto(0, {casoId}, {idAccion}, \'{nombre}\')" title="Agregar gasto"><img width="16px" src="../images/iconos/iconos_grid/agregar.png"></a>';
                // $link .= '&nbsp; <a class="grid_delete" href="javascript:void(0);" onclick="eliminarAccion({idAccion})" title="Eliminar acci&oacute;n"></a>';
                $column->ItemTemplate = $link;
            }
            elseif($grid->id == "clientes"){
                $link = '';
                $link .= '<a class="grid_edit" onclick="grid_edit(this)" href="javascript:void 0" title="Editar"></a>';
                $column->ItemTemplate = $link;
            }
            elseif ($grid->id == "cat_tipo_casos") {
                $link = '';
                $link .= '<a class="grid_edit" onclick="grid_edit(this)" href="javascript:void 0" title="Editar"></a>';
                $link .= '<a class="grid_delete" onclick="grid_delete(this)" href="javascript:void 0" title="Eliminar"></a>
                ';
                $column->ItemTemplate = $link;
            }
            elseif ($grid->id == "cat_conceptos") {
                $link = '';
                $link .= '<a class="grid_edit" onclick="grid_edit(this)" href="javascript:void 0" title="Editar"></a>';
                $link .= '<a class="grid_delete" onclick="grid_delete(this)" href="javascript:void 0" title="Eliminar"></a>';
                $column->ItemTemplate = $link;
            }elseif ($grid->id == "gastosGrid") {
                $link = '';
                $link .= '<a class="grid_list btnFancy" onclick="verTablaConceptos({idUsuario}, {totalAbonos}, {totalCargos}, {saldoPeriodo}, {saldo})" href="#fancyConcepto" title="Lista"></a>
                <a class="grid_abono btnFancy" onclick="reseteaFormulario(\'formAbono\');seleccionaAbogado({idUsuario}, \'_a\');" href="#fancyNuevoAbono" title="Agregar entrada"></a>
                <a class="grid_cargo btnFancy" onclick="reseteaFormulario(\'formCargo\');seleccionaAbogado({idUsuario}, \'_c\');" href="#fancyNuevoCargo" title="Agregar salida"></a>

                <form role="form" id="formImprimir_{idUsuario}" name="formImprimir_{idUsuario}" method="post" action="conceptos.php?idUsuario={idUsuario}" target="_blank" >
                <input type="hidden" name="totalAbonos_{idUsuario}" id="totalAbonos_{idUsuario}" value="{totalAbonos}">
                <input type="hidden" name="totalCargos_{idUsuario}" id="totalCargos_{idUsuario}" value="{totalCargos}">
                <input type="hidden" name="saldoPeriodo_{idUsuario}" id="saldoPeriodo_{idUsuario}" value="{saldoPeriodo}">
                <input type="hidden" name="saldo_{idUsuario}" id="saldo_{idUsuario}" value="{saldo}">
                <input type="hidden" name="desde_{idUsuario}" id="desde_{idUsuario}" value="">
                <input type="hidden" name="hasta_{idUsuario}" id="hasta_{idUsuario}" value="">
                <!-- <a class="grid_view" onclick="imprimirGastos({idUsuario}, 0);" title="Ver"></a> -->
                <a class="grid_print" onclick="imprimirGastos({idUsuario}, 1);" title="Imprimir"></a>
                </form>
                ';
                
                $column->ItemTemplate = $link;
            }
            elseif($grid->id == "grid_actividades"){
                $link = '';
               
                $link .= '<a class="grid_edit" target="_blank" href="actividad.php?expId={casoId}&actId={idAccion}" title="Ver detalle acci&oacute;n"></a>';
              
                $column->ItemTemplate = $link;
            }
            elseif($grid->id == "tareas"){
                $link = '';
               
                $link .= '<a class="grid_edit" target="_blank" href="tarea.php?tareaId={idTarea}" title="Ver detalle tarea"></a>';
               
                $column->ItemTemplate = $link;
            }
            elseif($grid->id == "grid_tareas"){
                $link = '';
               
                $link .= '<a class="grid_edit" target="_blank" href="tarea.php?tareaId={idTarea}" title="Ver detalle acci&oacute;n"></a>';
              
                $column->ItemTemplate = $link;
            }elseif ($grid->id=="digitales") {
                $link = '';
                $link .= '
                <a class="grid_edit" data-toggle="modal" data-target="#popup_modalDigital" title="Editar" onclick="edicionDigital({idDocumento}, {tipo}, \'{nombre}\', \'{url}\')"></a>
                <a class="grid_delete btnFancyBox" onclick="verificaUsoTabla(\'digitales\', {idDocumento})" href="#fancyElimCat" title="Eliminar"></a>
                <a class="grid_open"  onclick="descargaDigital(\'{url}\', {tipo})" title="Ver"></a>
                ';
                $column->ItemTemplate = $link;
            }
            elseif($grid->id == "pagos" || $grid->id == "adicionales"){
                $column->ItemTemplate = '
                <a class="{claseRecibo}" title="Ver recibo" target="_blank" href="../upload/recibos/{recibo}"></a>
                ';
            }
            elseif($grid->id == "cuentas"){
                $column->ItemTemplate = '
                <a class="grid_edit" href="cuentasxcobrar.php?expedienteId={casoId}" title="Editar" target="_blank"></a>
                ';
            }
            else{
                $column->ItemTemplate = '
                <a class="grid_edit" onclick="grid_edit(this)" href="javascript:void 0" title="Editar"></a>'
                . '<a class="grid_delete" onclick="grid_delete(this)" href="javascript:void 0" title="Eliminar"></a>'
                // . '<a class="btnDesactivarUsuario" onclick="muestraDesactivarUsuario({idUsuario},\'usuariosGrid\',\'usuarios\',{activo})"  href="#fancyDesactivarUsuario" title="Activar/Desactivar usuario"><img src="../images/{nombreImg}" class="iconoDesactivar" ></a>'
                ;
            }
        }elseif($_SESSION['idRol']==2){
              if ($grid->id=="comunicados") {
                $column->ItemTemplate = '
                <a class="grid_edit" onclick="grid_edit(this)" href="javascript:void 0" title="Editar"></a>'
                . '<a class="grid_delete" onclick="grid_delete(this)" href="javascript:void 0" title="Eliminar"></a>'
                // . '<a class="btnDesactivarUsuario" onclick="muestraDesactivarUsuario({idUsuario},\'usuariosGrid\',\'usuarios\',{activo})"  href="#fancyDesactivarUsuario" title="Activar/Desactivar usuario"><img src="../images/{nombreImg}" class="iconoDesactivar" ></a>'
                ;
              }elseif($grid->id == "comunicados"){
                $column->ItemTemplate = '<a class="grid_edit" onclick="grid_edit(this)" href="javascript:void 0" title="Editar"></a>';
                //<a class="grid_edit" onclick="edicionGrid(\'comunicados\', {idComunicado});" href="javascript:void 0" title="Editar"></a>'
                    //.' <a class="grid_delete" onclick="grid_delete(this)" href="javascript:void 0" title="Eliminar"></a>'
                // . '<a class="btnDesactivarUsuario" onclick="muestraDesactivarUsuario({idUsuario},\'usuariosGrid\',\'usuarios\',{activo})"  href="#fancyDesactivarUsuario" title="Activar/Desactivar usuario"><img src="../images/{nombreImg}" class="iconoDesactivar" ></a>'
                ;
            }
            else {
               $column->ItemTemplate = '
               <a class="grid_edit" onclick="grid_edit(this)" href="javascript:void 0" title="Editar"></a>'
               // . '<a class="btnDesactivarUsuario" onclick="muestraDesactivarUsuario({idUsuario},\'usuariosGrid\',\'usuarios\',{activo})"  href="#fancyDesactivarUsuario" title="Activar/Desactivar usuario"><img src="../images/{nombreImg}" class="iconoDesactivar" ></a>'
               ;
              }
        }elseif($_SESSION['idRol']==5){  //Rol clientes
            if($grid->id == "casos"){
                $link = '';
                // $link .= '<a href="frmExpedienteEdit.php?id={idCaso}" title="Ver"><img src="../images/iconos/iconos_grid/ver.png"></a>';
                $link .= '<a class="grid_edit" onclick="edicionGrid(\'casos\', {idCaso});" href="javascript:void 0" title="Editar"></a>
                ';
                $column->ItemTemplate = $link;
            }
            if($grid->id == "caso_acciones"){
                $link = '';
                $link .= '<a class="grid_edit" target="_blank" href="actividad.php?expId={casoId}&actId={idAccion}" title="Ver detalle acci&oacute;n"></a>';
                $column->ItemTemplate = $link;
            }
            if($grid->id == "tareas"){
                $link = '';
               
                $link .= '<a class="grid_edit" target="_blank" href="tarea.php?tareaId={idTarea}" title="Ver detalle tarea"></a>';
               
                $column->ItemTemplate = $link;
            }
        }elseif($_SESSION['idRol']==4){  //Rol abogados
            if($grid->id == "casos"){
                $link = '';
                $link .= '<a class="grid_edit" onclick="edicionGrid(\'casos\', {idCaso});" href="javascript:void 0" title="Editar"></a>';
                $column->ItemTemplate = $link;
            }
            if($grid->id == "caso_acciones"){
                $link = '';
                $link .= '<a class="grid_edit" target="_blank" href="actividad.php?expId={casoId}&actId={idAccion}" title="Ver detalle acci&oacute;n"></a>';
                // $link .= '&nbsp; <a href="javascript:void(0);" onclick="popupCreaEditaGasto(0, {casoId}, {idAccion}, \'{nombre}\')" title="Agregar gasto"><img width="16px" src="../images/iconos/iconos_grid/agregar.png"></a>';
                // $link .= '&nbsp; <a href="javascript:void(0);" onclick="eliminarAccion({idAccion})" title="Eliminar acci&oacute;n"><img width="16px" src="../images/iconos/iconos_grid/eliminar.png"></a>';
                $column->ItemTemplate = $link;
            }elseif ($grid->id == "gastosGrid") {
                $link = '';
                $link .= '<a class="grid_list btnFancy" onclick="verTablaConceptos({idUsuario}, {totalAbonos}, {totalCargos}, {saldoPeriodo}, {saldo})" href="#fancyConcepto" title="Lista"></a>
                <a class="grid_abono btnFancy" onclick="reseteaFormulario(\'formAbono\');seleccionaAbogado({idUsuario}, \'_a\');" href="#fancyNuevoAbono" title="Agregar entrada"></a>
                <a class="grid_cargo btnFancy" onclick="reseteaFormulario(\'formCargo\');seleccionaAbogado({idUsuario}, \'_c\');" href="#fancyNuevoCargo" title="Agregar salida"></a>
                
                <form role="form" id="formImprimir_{idUsuario}" name="formImprimir_{idUsuario}" method="post" action="conceptos.php?idUsuario={idUsuario}" target="_blank" >
                <input type="hidden" name="totalAbonos_{idUsuario}" id="totalAbonos_{idUsuario}" value="{totalAbonos}">
                <input type="hidden" name="totalCargos_{idUsuario}" id="totalCargos_{idUsuario}" value="{totalCargos}">
                <input type="hidden" name="saldoPeriodo_{idUsuario}" id="saldoPeriodo_{idUsuario}" value="{saldoPeriodo}">
                <input type="hidden" name="saldo_{idUsuario}" id="saldo_{idUsuario}" value="{saldo}">
                <input type="hidden" name="desde_{idUsuario}" id="desde_{idUsuario}" value="">
                <input type="hidden" name="hasta_{idUsuario}" id="hasta_{idUsuario}" value="">
                <a class="grid_print" onclick="imprimirGastos({idUsuario});" title="Imprimir"></a>
                </form>
                ';
                
                $column->ItemTemplate = $link;
            }
            elseif($grid->id == "grid_actividades"){
                $link = '';
               
                $link .= '<a class="grid_edit" target="_blank" href="actividad.php?expId={casoId}&actId={idAccion}" title="Ver detalle acci&oacute;n"></a>';
              
                $column->ItemTemplate = $link;
            }
            elseif($grid->id == "grid_tareas" || $grid->id == "tareas" ){
                $link = '';
               
                $link .= '<a class="grid_edit" target="_blank" href="tarea.php?tareaId={idTarea}" title="Ver detalle tarea"></a>';
               
                $column->ItemTemplate = $link;
            }elseif ($grid->id=="digitales") {
                $link = '';
                $link .= '
                <a class="grid_edit" data-toggle="modal" data-target="#popup_modalDigital" title="Editar" onclick="edicionDigital({idDocumento}, {tipo}, \'{nombre}\')"></a>
                <a class="grid_delete btnFancyBox" onclick="verificaUsoTabla(\'digitales\', {idDocumento})" href="#fancyElimCat" title="Eliminar"></a>
                <a class="grid_open" onclick="descargaDigital(\'{url}\', {tipo})" title="Ver" ></a>
                ';
                $column->ItemTemplate = $link;
            }
        }

        $column->Align = "center";
        $column->HeaderText = "Acciones";
        $column->Width = "80px";
        $grid->MasterTable->AddColumn($column);
    }
}