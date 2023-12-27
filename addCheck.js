/*Crear check*/
// Check Asignación
function onEdit(e){

    var rowDate = e.range.getRow();  
    var columnDate = e.range.getColumn();
    var rowDateS = e.range.getRow();  
    var columnDateS = e.range.getColumn();
    
  //Fecha Asignacion 
if(columnDate === 5 && rowDate >1  && e.source.getActiveSheet().getName() === "Carga de TK nuevos"){
    
e.source.getActiveSheet().getRange(rowDate,6).setValue(new Date());
    }

  //Actualizacion de sitios
if(columnDate === 5 && rowDate >=8 && rowDate <=14  && e.source.getActiveSheet().getName() === "Actualizacion de sitios"){
    
e.source.getActiveSheet().getRange(rowDate,7).setValue(new Date());
    }

if(columnDate === 9 && rowDate >=8 && rowDate <=14  && e.source.getActiveSheet().getName() === "Actualizacion de sitios"){
    
e.source.getActiveSheet().getRange(rowDate,11).setValue(new Date());
    }
if(columnDate === 9 && rowDate >=17 && rowDate <=17  && e.source.getActiveSheet().getName() === "Actualizacion de sitios"){
    
e.source.getActiveSheet().getRange(rowDate,11).setValue(new Date());
    }  

if(columnDate === 5 && rowDate >1  && e.source.getActiveSheet().getName() === "Stand by"){

e.source.getActiveSheet().getRange(rowDate,6).setValue(new Date());
    }
    }
