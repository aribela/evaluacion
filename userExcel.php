<?php
include_once 'brules/PHPExcel/PHPExcel.php';
$dirname = dirname(__DIR__);
include_once $dirname.'/brules/usuariosObj.php';
$dateByZone = new DateTime("now", new DateTimeZone('America/Mexico_City') );
$dateTime = $dateByZone->format('Y-m-d H:i:s'); //fecha Actual

$data = obtenerRegistros();

crearExcel($data, "Usuarios");

function crearExcel($data = array(), $tituloExcel = "Titulo"){
    set_time_limit(0);
        date_default_timezone_set('America/Mexico_City');



        if (PHP_SAPI == 'cli')
                die('Este archivo solo se puede ver desde un navegador web');

        // Se crea el objeto PHPExcel
        $objPHPExcel = new PHPExcel();

        // Se asignan las propiedades del libro
        $objPHPExcel->getProperties()->setCreator("Monzani") //Autor
                                    ->setLastModifiedBy("Monzani") //Ultimo usuario que lo modifico
                                    ->setTitle("Reportes")
                                    ->setSubject("Reportes")
                                    ->setDescription("Reportes")
                                    ->setKeywords("Reportes")
                                    ->setCategory("Reportes");

        $colT = 'A';
        $FilaPrimera = 1;
        $FilaTitulo = 2;
        $objPHPExcel->getActiveSheet()->setCellValue($colT.$FilaPrimera, $tituloExcel);
        $colFinalTit = 'K';
        $filaT = 3;//Fila de inicio
        $contRegAnt = 0;
        foreach ($data as $key => $item) {
            $colB = 'A';
            $titulosColumnas = $item['titulos'];
            $registros = $item['registros'];
            $campos = $item['campos'];
            $FilaTitulo = ($key == 0)?$FilaTitulo:$FilaTitulo+$contRegAnt+2;
            $filaT = $FilaTitulo+1;
            $totalTituloCol = count($titulosColumnas);
            $contRegAnt = count($item['registros']);
            //recorre las preguntas que en este caso son los titulos
            // $objPHPExcel->getActiveSheet()->setCellValue("A1", "Fecha");
            for($i=1; $i<=$totalTituloCol; $i++){
                $objPHPExcel->getActiveSheet()->setCellValue($colB.$FilaTitulo, $titulosColumnas[$i-1]);
                $colB++;
            }
            $objPHPExcel->getActiveSheet()->mergeCells("A1:".$colB."1");
            
            foreach(range('A','O') as $columnID){
                $objPHPExcel->getActiveSheet()->getColumnDimension($columnID)->setWidth(10);
            }
            
            $contLetra = "$colB"."2";
            $iIns=2;
            
            //Inicio - insertar contenido
            $colA = 'A';//definimos columna de inicio
            $contReg = count($registros);//Total registros
            $revReg = 1; //Para identificar cuando lleguemos al ultimo registro (totales)
            $filaInicialTabla = $filaT; //fila inicial
            
            //Recorremos los registros cada uno es una fila
        
            foreach ($registros as $registro) {
                $esfilaTotales = ($revReg == $contReg)?true:false;

                //recorremos cada campo cada uno es una columna
                foreach ($campos as $campo) {
                    $valor = $registro->{$campo["nombre"]};
                    if($esfilaTotales){
                        // // echo $valor." ";
                        // $valor = str_replace("%%", $filaInicialTabla, $valor);//Reemplazar en la formula, la fila inicial para la suma
                        // $valor = str_replace("%", ($filaT-1), $valor);//Reemplazar en la formula, la fila final para la suma
                        // $valor = ($campo["nombre"] != 'sucursal')?"=".$valor."":$valor;
                        // // $valor = "";
                        // // echo $valor."<br>";
                    }
                    elseif($campo["formula"] != ""){
                        $valor = str_replace("_", $filaTotalTabla, $valor);//Reemplazar en la formula la fila de total (si la tiene)
                        $valor = str_replace("%", $filaT, $valor);//Reemplazar en la formula la fila actual
                    }
                    
                    if($campo["formato"] == "moneda"){
                        
                        $objPHPExcel->getActiveSheet()
                        ->getStyle($colA.$filaT)
                        ->getNumberFormat()
                        ->setFormatCode(
                            '_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)'
                        );
                        $valor = str_replace("$", "", $valor);
                        $valor = str_replace(",", "", $valor);
                        

                        //Jair 7/Ago/2020 Quitar iva
                        // $valor = floatval($valor)/1.16;

                        $objPHPExcel->setActiveSheetIndex(0)->setCellValue($colA.$filaT,floatval($valor));
                        
                    }
                    elseif($campo["formato"] == "link"){
                        $objPHPExcel->setActiveSheetIndex(0)->setCellValue($colA.$filaT,$valor);
                        //saber si una cadena es una url
                        if(filter_var($valor, FILTER_VALIDATE_URL)){
                            // $valor = '=hyperlink("'.$valor.'";"Recibo")';
                            $objPHPExcel->setActiveSheetIndex(0)->setCellValue($colA.$filaT,$valor);
                            $objPHPExcel->getActiveSheet()->getCell($colA.$filaT)->getHyperlink()->setUrl($valor);
                        }


                    }
                    elseif($campo["formato"] == "porcentaje"){
                        $objPHPExcel->getActiveSheet()->getStyle($colA.$filaT)
                        ->getNumberFormat()->applyFromArray( 
                            array( 
                                'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00
                            )
                        );
                        // echo "valor: ".$valor."  fila: ".$colA." ".$filaT." ".$campo["nombre"]."<br>";
                        // echo "<pre>";print_r($valor);echo "</pre>";
                        $valor = ($valor == "")?0:$valor;
                        $objPHPExcel->setActiveSheetIndex(0)->setCellValue($colA.$filaT,$valor/100);
                    }
                    elseif($campo["formato"] == "subfilas"){
                        $arrPagos = explode(",", $valor);
                        if(count($arrPagos) > 0){
                            foreach ($arrPagos as $pago) {
                                if($pago != ''){
                                    $colA = 'A';
                                    $filaT++;
                                    $datosPago = explode("|", $pago);
                                    $contSubCol = 1;
                                    foreach ($datosPago as $datoPago) {
                                        if($contSubCol == 1){//columna uno subfila - monto pago
                                            $objPHPExcel->getActiveSheet()
                                            ->getStyle('I'.$filaT)
                                            ->getNumberFormat()
                                            ->setFormatCode(
                                                '_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)'
                                            );
                                            $datoPago = str_replace("$", "", $datoPago);
                                            $datoPago = str_replace(",", "", $datoPago);
                                            $datoPago = -1*floatval($datoPago);
        
                                            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('I'.$filaT,$datoPago);
                                        }else{// fecha pago
                                            $fechaPago = ($datoPago != '' && strlen($datoPago) >= 10)?convertirFechaVistaConHora($datoPago):'';
                                            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('H'.$filaT, $fechaPago);
                                        }
                                        $colA++;
                                        $contSubCol++;
                                    }
                                }
                            }
                        }

                    }
                    else{
                        $objPHPExcel->setActiveSheetIndex(0)->setCellValue($colA.$filaT, html_entity_decode($valor));
                    }
                    
                    $colA++;

                    //Old way
                    //$objPHPExcel->setActiveSheetIndex(0)->setCellValue($colA.$filaActual,html_entity_decode($registro->$campo));
                    //$colA++;
                }
                $colA = 'A';
                $filaT++;
                $revReg++;//Aumentar el contador para saber si vamos en el ultimo registro

            }//Fin de inserción de filas
            // die();
            //<editor-fold defaultstate="collapsed" desc="Estilos columnas" >
            $estiloTituloColumnas = array('font' => array('name'=> 'Arial',
            'size'=>8,
            'bold'=> true,
            'color'=> array('rgb' => '000000')
            ),
            'alignment' => array('horizontal'=>PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            'vertical'=>PHPExcel_Style_Alignment::VERTICAL_CENTER,
            'wrap'=>TRUE
            ),
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => '69d4cc'))
            );
            $objPHPExcel->getActiveSheet()->getStyle('A'.$FilaTitulo.':'.$colFinalTit.$FilaTitulo)->applyFromArray($estiloTituloColumnas);
        }//fin for data   /// ***************************************************************************************

        $estiloInformacion = new PHPExcel_Style();
        $estiloInformacion->applyFromArray(array('font' => array('name'=>'Arial',
                                'size'=>8,
                                'color'=> array('rgb' => '000000')
                            )
            )
        );
        //</editor-fold>

        //$objPHPExcel->getActiveSheet()->setSharedStyle($estiloInformacion, 'A3:AZ3'.($i-1));

        for($j = 'A'; $j <=$contLetra; $j++){
            $objPHPExcel->setActiveSheetIndex(0)->getColumnDimension($j)->setAutoSize(TRUE);
        }


        // Se asigna el nombre a la hoja
        $objPHPExcel->getActiveSheet()->setTitle($nombreRep);

        // Se activa la hoja para que sea la que se muestre cuando el archivo se abre
        $objPHPExcel->setActiveSheetIndex(0);
        // Inmovilizar paneles
        $objPHPExcel->getActiveSheet(0)->freezePaneByColumnAndRow(0,1);

        //Mostrar notas en las columnas que lo requieran en una nueva hoja
        $filaI = 1;
        if(count($colNotas) > 0){
            $objPHPExcel->createSheet();
            foreach($colNotas as $colnota){
                //col
                //nombre
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue("A".$filaI, $colnota["col"]);
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue("B".$filaI, $colnota["nombre"]);
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue("C".$filaI, $colnota["nota"]);
                $filaI++;
            }
        }

        $titletmp = str_replace(" ", "_", $nombreRep).'_'."_".date("d-m-Y");
        $titleFile = "Reporte_".$titletmp.".xls";
        header('Content-Type: application/vnd.ms-excel');
        header("Content-Disposition: attachment;filename=$titleFile");
        header('Cache-Control: max-age=0');

        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save('php://output');
        exit;
        ob_flush();
}

function obtenerRegistros(){
    $arrTitulos = array(
        "Metodo",
        "Total"
    );
    $arrRegistros = array();

    $usuariosObj = new usuariosObj();

    $usuarios = $usuariosObj->obtTodosUsuarios();

    foreach ($usuarios as $usuario) {
        $arrRegistros[] = (object)array(
            "Login" => $usuario->Login,
            "NombreCompleto" => $usuario->Nombres.' '.$usuario->Paterno.' '.$usuario->Materno,
            "Sueldo" => $usuario->Sueldo,
            "FechaIngreso" => $usuario->FechaIngreso,
        );
    }


    $campos = array(
        array("nombre" => 'Login', "formula" => "", "formato" => ""),
        array("nombre" => 'NombreCompleto', "formula" => "", "formato" => ""),
        array("nombre" => 'Sueldo', "formula" => "", "formato" => "moneda"),
        array("nombre" => 'FechaIngreso', "formula" => "", "formato" => ""),
    );

    $arrData[] = array(
        "titulos" => $arrTitulos,
        "registros" => $arrRegistros,
        "campos" => $campos,
    );
// echo "<pre>";print_r($arrMetodos);echo "</pre>";die();
    return $arrData;
}