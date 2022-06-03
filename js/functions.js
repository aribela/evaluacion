/*
 * Date: 5/05/2017
 * Actualizado: 07/12/2020
 * Funciones generales y especificas de javascript
 */

$(document).ready(function(){     
});

function verPrimeros(limit){
    var params = {funct: 'verPrimeros', limit: limit};
  ajaxData(params, function(data){
    $('#contenido').html(data.html);
  });
}

function actualizaSueldo(userId, sueldo){
    var params = {funct: 'actualizaSueldo', userId: userId, sueldo: sueldo};
  ajaxData(params, function(data){
    alert("Sueldo actualizado");
  });
}

function verNuevo(){
    let html = '<div class="row">';
    html += '<div class="col-md-12">';
    html += '<div class="card">';
    html += '<div class="card-header">';
    html += '<h4 class="card-title">Nuevo usuario</h4>';
    html += '</div>';
    html += '<div class="card-body">';
    html += '<form id="formNuevo">';
    html += '<div class="row">';
    html += '<div class="col-md-6">';
    html += '<div class="form-group">';
    html += '<label>Nombre</label>';
    html += '<input type="text" class="form-control" name="nombre" placeholder="Nombre">';
    html += '</div>';
    html += '</div>';
    html += '<div class="col-md-6">';
    html += '<div class="form-group">';
    html += '<label>Paterno</label>';
    html += '<input type="text" class="form-control" name="Paterno" placeholder="Paterno">';
    html += '</div>';
    html += '</div>';
    html += '<div class="col-md-6">';
    html += '<div class="form-group">';
    html += '<label>Materno</label>';
    html += '<input type="text" class="form-control" name="Materno" placeholder="Materno">';
    html += '</div>';
    html += '</div>';
    html += '</div>';
    html += '<div class="row">';
    html += '<div class="col-md-6">';
    html += '<div class="form-group">';
    html += '<label>Sueldo</label>';
    html += '<input type="text" class="form-control" name="Sueldo" placeholder="Sueldo">';
    html += '</div>';
    html += '</div>';
    html += '<button type="button" class="btn btn-info btn-fill pull-right" onclick="guardar()">Guardar</button>';
    $('#contenido').html(html);
}

function guardar(){
    var formElement = document.getElementById("formNuevo");
    var params = new FormData(formElement);
    // params['funct'] = 'crearCaso';
    
    var urlDel = 'ajaxcall/ajaxFunctions.php?funct=guardar';
    ajaxDataPost(urlDel, params, function(data){
    // ajaxData(params, function(data){
      console.log(data);

      if(data.success){
        alert("cambios guardados");
      }
    });
}