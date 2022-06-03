/************************Inicializar plugin***************************************/	  
var gralClic = false;
// target elements with the "draggable" class
interact('.draggable')
  .draggable({
	// enable inertial throwing
	inertia: true,
	// keep the element within the area of it's parent
	restrict: {
	  restriction: "parent",
	  endOnly: true,
	  elementRect: { top: 0, left: 0, bottom: 1, right: 1 }
	},
	// enable autoScroll
	autoScroll: true,

	// call this function on every dragmove event
	onmove: dragMoveListener,
	// call this function on every dragend event
	onend: function (event) {
	  // console.log(event.target.id);
	  var idElemento = event.target.id;
	  idElementoArr = idElemento.split("_");
	  var ctr_hostpot = idElementoArr[1];
	  // edicionZonas(ctr_hostpot, 2);

	  $("#spanTituloZona").html("Editar");
	  $("#opc_version_zona").val(2);
		$("#idpunto_version_zona").val(ctr_hostpot);
		$("#ip_galeria").val($("#"+idElemento).attr("data-galeriaid"));
		$("#btnAgregarZona").click();

	},
	onstart: function(event){
		// console.log("start move");
		var idElemento = event.target.id;
		$("#x_version_zona").val( $("#"+idElemento).attr("data-x") );
		$("#y_version_zona").val( $("#"+idElemento).attr("data-y") );
	}
  });

  function dragMoveListener (event) {
	//console.log(event.target);		
	var target = event.target,
		// keep the dragged position in the data-x/data-y attributes
		x = (parseFloat(target.getAttribute('data-x')) || 0) + event.dx,
		y = (parseFloat(target.getAttribute('data-y')) || 0) + event.dy;

	// translate the element
	target.style.webkitTransform =
	target.style.transform =
	  'translate(' + x + 'px, ' + y + 'px)';

	// update the posiion attributes
	target.setAttribute('data-x', x);
	target.setAttribute('data-y', y);
	gralClic=true;
  }
	 
  // this is used later in the resizing and gesture demos
  window.dragMoveListener = dragMoveListener;	 
	  

/***********************Modificar init****************************************/
$( document ).ready(function() {
	$("#cont_hostpot").css("width",$("#anchoImg").val());
	$("#cont_hostpot").css("height",$("#altoImg").val());

	var altomenu = 150+$("#altoImg").val();
	$("#menu2").css("height", 500);

	$('#cont_hostpot').dblclick(function (e) {
		ctr_hostpot = parseInt($("#ctr_hostpot").val())+1;
		xHostpost = e.pageX-470;
		yHostpost = e.pageY-585;
		console.log(xHostpost);
		console.log(yHostpost);
		hostpot = '';
		hostpot = '<div class="draggable hostpot" onclick="clickHostpot('+ctr_hostpot+',2);" data-x="'+xHostpost+'" data-y="'+yHostpost+'" data-activazonaid="0" data-galeriaid="0" style="transform: translate('+xHostpost+'px, '+yHostpost+'px);" id="hostpot_'+ctr_hostpot+'"><p></p></div>';			
		$("#cont_hostpot").append(hostpot);			
		$("#ctr_hostpot").val(ctr_hostpot);
		$("#spanTituloZona").html("Agregar");

		$("#opc_version_zona").val(1);
		$("#idpunto_version_zona").val(ctr_hostpot);

		versionZonaGrid.refresh();
        versionZonaGrid.commit();
		$("#btnAgregarZona").click();
		// clickHostpot(ctr_hostpot, 1);
	});


	if($("#ctr_hostpot").val() > 0)
	{
		var ctr_hostpot = 1;
		$("input.punto_hidden").each(function(){
			var coordenada = $(this).val();
			var coordenadas = coordenada.split(",");
			xHostpost = coordenadas[0];
			yHostpost = coordenadas[1];


			var idPunto = this.id;
			var idPuntoArr = idPunto.split("_");
			var activaZonaId = idPuntoArr[1];

			var galeriaId = $("#"+idPunto).attr("data-galeriaid");

			console.log(xHostpost);
			console.log(yHostpost);
			var hostpot = '';
			hostpot = '<div class="draggable hostpot" onclick="clickHostpot('+ctr_hostpot+', 2);" data-x="'+xHostpost+'" data-y="'+yHostpost+'" data-activazonaid="'+activaZonaId+'" data-galeriaid="'+galeriaId+'" style="transform: translate('+xHostpost+'px, '+yHostpost+'px);" id="hostpot_'+ctr_hostpot+'"></div>';
			$("#cont_hostpot").append(hostpot);
			$("#ctr_hostpot").val(ctr_hostpot);
			ctr_hostpot++;
        });
	}
});
//ejecuta al presionar sobre un hostpot
//opc: 1=agregar, 2=editar
function clickHostpot(ctr_hostpot, opc){
	if(gralClic==false){
		setTimeout(function(){			
	var activazonaid = $("#hostpot_"+ctr_hostpot).attr("data-activazonaid");

	window.open('galeriamultiple.php?activazonaid='+activazonaid,'_blank');
		}, 300);
	}

	// var mensaje = '<strong>&#191Esta seguro de agregar esta zona?</strong>';
	// mensaje += '<p><b>Color: </b>'+$("#ip_colores option:selected").text();+'</p>';
	// mensaje += '<p><b>Galeria: </b>'+$("#ip_galeria option:selected").text();+'</p>';

	// //Agregar zona
	// if(opc == 1)
	// {
	// 	alertify.confirm(mensaje, function(){
 //            edicionZonas(ctr_hostpot, opc);
 //        },function(){
 //          $("#hostpot_"+ctr_hostpot).remove();
 //        }).set({labels:{ok:'Aceptar', cancel: 'Cancelar'}, padding: false});
	// }//Mover zona
	// else
	// {
	// 	// edicionZonas(ctr_hostpot, opc);
	// }
	 
}


function edicionZonas(ctr_hostpot, opc)
{
	// console.log($("#hostpot_"+ctr_hostpot).data("x")+" - "+$("#hostpot_"+ctr_hostpot).data("y"));
	// console.log($("#hostpot_"+ctr_hostpot).attr("data-x"));
	 params = {
        funct: 'edicionZonas',
        opc: opc,
        coloresVersId: $("#ip_colores").val(),
        galeriaId: $("#ip_galeria").val(),
        coordenadas: $("#hostpot_"+ctr_hostpot).attr("data-x")+","+$("#hostpot_"+ctr_hostpot).attr("data-y"),
        activaZonaId: $("#hostpot_"+ctr_hostpot).attr("data-activazonaid"),
    };
     console.log("antes")       ;
    ajaxData(params, function(data){
    	console.log(data);
    	if(data.success)
    	{
    		parent.$.fancybox.close();
    		if(data.existe)
    		{
    			alertify.error("Ya existe una zona para el color y galeria seleccionados");
    			$("#hostpot_"+ctr_hostpot).remove();
    		}

    		if(data.res > 0)
    		{
    			alertify.success("Zona agregada correctamente");
    			$("#hostpot_"+ctr_hostpot).attr("data-activazonaid",data.res);
    			$("#hostpot_"+ctr_hostpot).attr("data-galeriaid",data.galeriaId);
    			versionZonaGrid.refresh();
                versionZonaGrid.commit();
    			if($("#btnMenu6").hasClass("tabdisabled"))
    			{
    				location.reload();
    			}
    		}

    		if(data.update > 0)
    		{
    			alertify.success("Ubicacion actualizada");
    		}
    			
    	}
    	else
    	{
    		alertify.error("Error inesperado, intente mas tarde");
    	}
    });
}

function mostrarFancyAgregarZona()
{

}

function cancelarAgregarZona()
{
	var ctr_hostpot = $("#idpunto_version_zona").val();
	if($("#opc_version_zona").val() == 1)
	{
		$("#hostpot_"+ctr_hostpot).remove();
	}
	if($("#opc_version_zona").val() == 2)
	{
		var x = $("#x_version_zona").val();
		var y = $("#y_version_zona").val();

		// var target = document.getElementById("#hostpot_"+ctr_hostpot);
		// translate the element
		// target.style.webkitTransform =
		// target.style.transform =
		//   'translate(' + x + 'px, ' + y + 'px)';

		  $("#hostpot_"+ctr_hostpot).css("webkitTransform",'translate(' + x + 'px, ' + y + 'px)');
		  $("#hostpot_"+ctr_hostpot).css("transform",'translate(' + x + 'px, ' + y + 'px)');

		// update the posiion attributes
		// target.setAttribute('data-x', x);
		// target.setAttribute('data-y', y);
		 $("#hostpot_"+ctr_hostpot).attr('data-x', x);
		 $("#hostpot_"+ctr_hostpot).attr('data-y', y);
		 $("#x_version_zona").val("");
		 $("#y_version_zona").val("");
	}
	
	parent.$.fancybox.close();
}


function guardarVersionZona()
{
	var ctr_hostpot = $("#idpunto_version_zona").val(); 
	var opc = $("#opc_version_zona").val();
	edicionZonas(ctr_hostpot, opc);
}