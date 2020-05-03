var express = require('express');
const excel = require('node-excel-export');
var app = express();
var dataset=[]
var SheeList=[] 

 const styles = {
    headerDark: {
      fill: {
        fgColor: {
          rgb: 'FF000000'
        }
      },
      font: {
        color: {
          rgb: 'FFFFFFFF'
        },
        sz: 14,
        bold: true,
        underline: true
      }
    },
    cellPink: {
      fill: {
        fgColor: {
          rgb: 'FFFFCCFF'
        }
      }
    },
    cellGreen: {
      fill: {
        fgColor: {
          rgb: 'FF00FF00'
        }
      }
    }
  };
 
app.get('/', function(req, res,next){   

      const specification = {
        SL: { 
           displayName: 'SL',
          headerStyle: styles.headerDark, // <- Header style
          cellStyle: styles.cellPink,
           width: 120 // <- width in pixels
         },
         id: {
           displayName: 'id',
           headerStyle: styles.headerDark,
           cellStyle: styles.cellPink, // <- Cell style
           width: '10' // <- width in chars (when the number is passed as string)
         },
         displayName: {
           displayName: 'displayName',
           headerStyle: styles.headerDark,
           cellStyle: styles.cellPink, // <- Cell style
          width: 220 // <- width in pixels
         },

         originalarrivaltime: {
            displayName: 'originalarrivaltime',
            headerStyle: styles.headerDark,
            cellStyle: styles.cellPink, // <- Cell style
           width: 220 // <- width in pixels
          },
          content: {
            displayName: 'content',
            headerStyle: styles.headerDark,
            cellStyle: styles.cellPink, // <- Cell style
           width: 220 // <- width in pixels
          }
        
       }

const fs = require('fs');
     fs.readFile('skype-messages.json', (err, data) => {
     //fs.readFile('test.json', (err, data) => {
     if (err) throw err;
     let skypemessages = JSON.parse(data);
     var PreviousGroup="";
     for(var i=0; i<  skypemessages.length;i++){

         console.log( skypemessages[i].group);       
         if(PreviousGroup!=skypemessages[i].group)
         {

            dataset=[];
             PreviousGroup=skypemessages[i].group;
             for(var j = 0; j < skypemessages[i].messages.length; j++)
             {
                var originalarrivaltime = skypemessages[i].messages[j].originalarrivaltime              
                var  PreviousDate=""
                 if(originalarrivaltime.substring(0, 8)=="2020-04-")
                 {
                     var dateofMonth=originalarrivaltime.substring(0, 10)
                     dataset.push({ 'SL' :j+1, 'id' : skypemessages[i].messages[j].id,'displayName':skypemessages[i].messages[j].displayName ,'originalarrivaltime' :dateofMonth,'content':skypemessages[i].messages[j].content});
                    console.log(j+1)
                 } 
             }
            SheeList.push({
                name: skypemessages[i].group    , 
                specification: specification, 
                data: dataset 
            })  
        }
     } 

console.log(dataset);

    const report = excel.buildExport(  
        SheeList
      );

      res.attachment('report.xlsx'); 
      return res.send(report);
    });
    
});
app.listen(3000);
console.log('Listening on port 3000');





