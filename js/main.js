var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet0 = ss.getSheets()[0];
var sheetdatabase = ss.getSheets()[1];
var ydict = {jan : '01', feb : '02', mar : '03', apr : '04', may : '05', jun : '06',jul : '07', aug : '08', sep : '09', oct : '10', nov : '11', dec : '12'};
var failed = 0, counted = 0;

function doGet() {
    failed = 0;
	var num = (sheet0.getLastRow()).toFixed(0).toString();
	var values = sheet0.getDataRange().getValues();
    Logger.log(num);
    for(var i = 1; i < num; i++ ){
        parseRSS(values[i][0].toString(), values[i][1].toString(), values[i][2].toString(), i+1); 
    }
	RSSdatabase(1);
}


function parseRSS(Name, feed, puD, count) {
   
  	var item, date, title, link, desc, encl, tag; 
  	var ArrayTemp = new Array();
    var txt;
  	var lastdate = puD;
    if ((lastdate!=null) && (lastdate!='')){
		try{
        	var result = UrlFetchApp.fetch(feed, { muteHttpExceptions: true });
        
 			Logger.log("code: " + result.getResponseCode());
 			Logger.log("text: " + result.getContentText());
        	txt = result.getContentText();
      }catch(e){
        
        var currentdate = new Date(); 
        var eventtime = currentdate.getFullYear() + "/"
                     + (currentdate.getMonth()+1)  + "/" 
                     + currentdate.getDate() + " "
                     + currentdate.getHours() + ":"  
                     + currentdate.getMinutes() + ":" 
                     + currentdate.getSeconds() + " UTC+0800";
        
        
        sheet0.getRange("D"+count).setValue(e);
        sheet0.getRange("E"+count).setValue(eventtime);
      }
      try {
        //var txt = loadXMLDoc(feed);   
    	var doc = Xml.parse(txt, false);   
    	var items = doc.getElement().getElement("channel").getElements("item"); 

    	//loop for add items
    	for (var i in items) {     
    		item  = items[i];      
    		date  = item.getElement("pubDate").getText().replace(/([A-Za-z]{3}\, )/ig,'');   
    		
    		//date split
    		ArrayTemp = date.split(" ",5);
    		
    		//prevent single number
      		if(ArrayTemp[0].length < 2)
        		ArrayTemp[0] = "0" + ArrayTemp[0];
    		
    		//rebuild date format
      		date  = ArrayTemp[2] + "/" +  
        	  ydict[ArrayTemp[1].toLowerCase()] + "/" + 
                    ArrayTemp[0] + " " + ArrayTemp[3] + " UTC" + ArrayTemp[4];
      		
    		//check last update
      		if(date==lastdate){failed=0;break;}
      		//update the what time is update
    		if(i==0) {sheet0.getRange("C"+count).setValue(date);}

    		//title
      		title = LanguageApp.translate(item.getElement("title").getText(),"zh-CN","zh-TW");
      		//link
      		link  = item.getElement("link").getText();
      		desc  = item.getElement("description").getText().replace('<![CDATA['.'').replace(']]>'.'');
      		//magnet
      		encl  = item.getElement("enclosure").getAttribute('url').getValue();
            tag   = LanguageApp.translate(item.getElement("category").getText(),"zh-CN","zh-TW").replace("ＲＡＷ","RAW");

      		//item add to database
      		sheetdatabase.appendRow([Name, title, desc, link, encl, date, tag]);
        }
        var currentdate = new Date(); 
        var eventtime = currentdate.getFullYear() + "/"
                  + (currentdate.getMonth()+1)  + "/" 
                  + currentdate.getDate() + " "
                  + currentdate.getHours() + ":"  
                  + currentdate.getMinutes() + ":" 
                  + currentdate.getSeconds() + " UTC+0800";
      
        sheet0.getRange("F"+count).setValue(eventtime);
      } catch(e) {

      }
    }

}

function RSSdatabase(count){
  var sheet = ss.getSheets()[count];
  
  var columnToSortBy = 5;
  var num = (sheet.getLastRow()).toFixed(0).toString();
  var range = sheet.getRange("A2:H"+num);
  range.sort( { column : columnToSortBy, ascending: false } );
}

function SameClear(){

  var sheet = ss.getSheets()[1];
  var range = sheet.getRange("C2:C100");
  var values = range.getValues();
  var lastvalue = '',temp;
  var j=2;
  for ( var i in values){
    temp = values[i].toString();
    if (temp == lastvalue){
      sheet.deleteRow(j);
    }else{lastvalue=values[i].toString();j++;}
  }
  
}
