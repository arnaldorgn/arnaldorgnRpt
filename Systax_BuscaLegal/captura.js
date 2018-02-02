var page = new WebPage()
var fs = require('fs');
var system = require('system');
var address = system.args[1]; 
var pathFileToSave = system.args[2]; 

function getTimestamp(){
    // or use Date.now()
    return new Date().getTime();
}

  
if (address == 'help') {
	console.log("\n\n\n");
	console.log('************************************************');
	console.log('*************Captura de Pagina *****************');
	console.log('************************************************');
	
	console.log('Modo de uso');
	console.log('Parametro 1 - obrigatorio - Url desejada');
	console.log('Parametro 2 - output - opcional - Caminho onde o arquivo deve ser salvo / print');
	console.log('obs: verificar se o diretorio tem permissao de escrita.. nao utilizar C: ');
	console.log("\n\n"+'Exemplo de uso: phantomjs.exe captura.js http://acesso.mte.gov.br/legislacao/atos-declaratorios.htm C:\\Users\\PC\\Documents\\');
	console.log("\n"+'************************************************');
	console.log('************************************************');
	console.log('************************************************');
	phantom.exit();
}
page.onResourceRequested = function(request) {
    // update the timestamp when there is a request
    lastTimestamp = getTimestamp();
};

page.onResourceReceived = function(response) {
    // update the timestamp when there is a response
    lastTimestamp = getTimestamp();
};


page.open(address, function(status) {
	
	console.log("Status: " + status);
	if(status === "success") {
		console.log('pagina encontrada processando...');
		
	} else {
		console.log('falha ao carregar o endereco');
		phantom.exit(1);
	}
});
	
	
	
page.onLoadFinished = function() {
  //console.log("page load finished");
  //page.render('export.png');
  var fileName='';
  
  if (pathFileToSave != undefined) {
	  fileName= pathFileToSave;
  } else {
	  fileName= fs.workingDirectory + "/" ;
  }

  if (fileName != 'print') {
	  if (fs.isWritable(fileName)) {
		  

		  var partsAdress = address.split("/");
		  fileName += partsAdress[partsAdress.length - 1];
		  
		  console.log("Arquivo Salvo: " + fileName);
		  
		  fs.write( fileName+'.html', page.content, 'w');
		  
		  console.log('Captura Finalizada');
	  } else {
		  console.log('sem permissao de escrita: '+ fileName);
	  }
  } else {
	  console.log( page.content);
  }
  
  phantom.exit();
};

/*
function checkReadyState() {
    setTimeout(function () {
		
        var curentTimestamp = getTimestamp();
        if(curentTimestamp-lastTimestamp>2000){
			console.log('Sem resposta fechando programa...');
			
            // exit if there isn't request or response in 1000ms
            phantom.exit();
        }
        else{
            checkReadyState();
        }
    }, 100);
}

checkReadyState();*/