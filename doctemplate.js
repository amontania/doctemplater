const fs = require('fs');
const axios = require("axios");
const PizZip =  require('pizzip');
const Docxtemplater= require('docxtemplater');
const content = fs.readFileSync('template2.docx', 'binary');
const ImageModule = require('docxtemplater-image-module-free');
const sizeOf = require('image-size')

//Below the options that will be passed to ImageModule instance


var opts = {}
opts.centered = false; //Set to true to always center images
opts.fileType = "docx"; //Or pptx
 
//Pass your image loader
opts.getImage = function(tagValue, tagName) {
    //tagValue is 'examples/image.png'
    //tagName is 'image'
    return fs.readFileSync(tagValue);
}
 
//Pass the function that return image size
opts.getSize = function(img, tagValue, tagName) {
    //img is the image returned by opts.getImage()
    //tagValue is 'examples/image.png'
    //tagName is 'image'
    //tip: you can use node module 'image-size' here
    const dimensions = sizeOf(img);

    return [dimensions.width, dimensions.height];
}
 
var imageModule = new ImageModule(opts);

const zip = new PizZip(content);
const doc = new Docxtemplater();
doc.loadZip(zip).attachModule(imageModule)


const data = fs.readFileSync('./notas.txt').toString();
let notasArray = [];
let tempranking =[];

// Get data from txt
let testData ={};
notasArray= data.replace(/\r?\n|\nr/g," ").split(" ");

for(let i=0; i < notasArray.length; i++){
    const nombre = notasArray[i].substring(0,notasArray[i].indexOf(":"));
    const nota   = notasArray[i].substring(notasArray[i].indexOf(":")+1, notasArray[i].length);
    const resultado = nota > 4 ? 'Aprobado':'Reprobado';
    let json = 
        {
            "nombre": nombre,
            "nota":  nota,
            "resultado":  resultado
        }
    

        tempranking.push(json);
   
}
 

// Get Aanother data
let products=  [
{ name: 'Cable', price: '25000', type : 'metro' },
{ name: 'Transformadores', price: '20000000', type : 'unitario' }
];



// Get data from Api external
let data3={};
  
	try {
		 axios.get("https://loripsum.net/api/10/short/headers/plaintext/decorate")
        .then((response) => {
            data3 =  response.data;
            sendword(data3);

            
           
        });
     
	} catch (err) {
       console.log(err);

	}
 





  function sendword(datos){  

 testData.title = "Administración Nacional de Electricidad";
 testData.alumnos =tempranking;
 testData.imglogo='./ande.png';
 testData.imgcontrato='./natura.jpg';
 testData.apertura=datos;
 testData.products= products;


 doc.setData(    testData     );

 try {
          doc.render();
     }catch(error){

         throw error;
     }
     let filename= "generado";
     const buf = doc.getZip().generate({type : 'nodebuffer'});
     fs.writeFileSync(`${filename}.docx`, buf);
   
    }


