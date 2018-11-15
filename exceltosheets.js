const xlsxj = require("xlsx-to-json-lc");
const bodyParser = require('body-parser');
const mongoose = require('mongoose');

const tabs = ["Sites", "Antennas" , "Antenna_Electrical_Parameters", "Sectors" , "Sector_Antenna_Ports", "Sector_Antennas"];

const collectlist= ["planet_sites","planet_antennas","planet_antEParams","planet_sectors","planet_secantennaports","planet_secantennas"]

console.time('Time Taken to Update Planet to Mongo');
readexceltojson(0);


function readexceltojson(it){
    console.time('Time Taken to Import '+ tabs[it]);
    xlsxj({        
        input: "../Read Planet/Read Tabs/Final_20181102.xlsx", 
        output: "output.json",
        sheet: tabs[it]
      }, function(err, result) {
        if(err) {
          console.error(err);
        }else {
            console.log("JSON object created")
            if(it==0)
            {
                excelltomongozero(result,it);
            }
            else if(it==1)
            {
                excelltomongoone(result,it);
            }
            else if(it==2)
            {
                excelltomongotwo(result,it);
            }
            else if(it==3)
            {
                excelltomongothree(result,it);
            }
            else if(it==4)
            {
                excelltomongofour(result,it);

            }
            else if (it==5)
            {
                excelltomongofive(result,it);
                
            }
        
       //console.log(result)
         
        }
      });
    }


const sheet_sites = new mongoose.Schema({
    "Site ID" : String,
    "Site UID" : String,
    "Site Name" :String,
    "Site Name 2" : String,
    "Candidate Priority" : String,
    Longitude	: Number,
     Latitude : Number,
     Description	: String,
     })

const sheet_antennas = new mongoose.Schema({
    "Site ID" : String,
    "Antenna ID": String,
    "Longitude" : Number,
    "Latitude":  Number,
    "Antenna File": String,
    "Height (m)" : Number,
    "Azimuth" : Number,
    "Mechanical Tilt" : Number,
    "Twist" : Number,
    "Donor Antenna" : String,
    "Terrain Height (m)" : Number,
    "Sectors" : String,
    "Number of Corrections"  : Number,
    "Correction (dB) at -180 degrees"  : Number,
    "Correction (dB) at -150 degrees"  : Number,
    "Correction (dB) at -120 degrees" : Number,
    "Correction (dB) at -90 degrees" : Number,
    "Correction (dB) at -60 degrees": Number,
    "Correction (dB) at -30 degrees": Number,
    "Correction (dB) at 0 degrees": Number,
    "Correction (dB) at 30 degrees": Number,
    "Correction (dB) at 60 degrees": Number,
    "Correction (dB) at 90 degrees": Number,
    "Correction (dB) at 120 degrees": Number,
    "Correction (dB) at 150 degrees": Number
})

const sheet_antEParams = new mongoose.Schema({
"Site ID" : String,
"Antenna ID" : Number,
"Electrical Controller" : String,
"Electrical Tilt" : Number,
"Electrical Azimuth" : Number,
"Electrical Beamwidth" : Number,
})

const sheet_sectors = new mongoose.Schema({    
    "Site ID" : String,
    "BTS Name" : String,
    "Sector ID" : String,
    "Sector UID": String,
    "Cell ID" : Number,
    "Technology": String,
    "Band Name": String,
    "Antenna Algorithm": String,
    "Propagation Model": String,
    "Distance (km)" : Number,
    "Radials" : Number,
    "Prediction Mode": String,
    "Interpolation Distance (m)" : Number,
    "Display Scheme": String,
    "Flag: Band": String,
    "Flag: BTS_Type": String,
    "Flag: Cabin_Type": String,
    "Flag: Data_Accuracy": String,
    "Flag: OP_Region": String,
    "Flag: Phase": String,
    "Flag: Property_Status": String,
    "Flag: Site_Status": String,
    "Flag: Tower_Owner": String,
    "Flag: Tower_Type": String,
    "Flag: Vendor": String
})


const sheet_secAntennas = new mongoose.Schema({   
    "Site ID" : String,
    "Sector ID": String,
    "Antenna ID": Number,
    "MIMO Group": String,
    "Horizontal Beamwidth": String,
    "Vertical Beamwidth": String,
    "Link Configuration ID" : String,
    "Cable Length (m)" : Number,
    "Power Split (%)" : Number,
    "Downlink Beamforming Configuration": String,
    "Uplink Beamforming Configuration": String
})

const sheet_secAntPorts = new mongoose.Schema({  
    "Site ID": String,
    "Sector ID": String,
    "Antenna ID" : Number,
    "Port Name": String,
    "Downlink": String,
    "Uplink": String
})
const schemalist = [sheet_sites,sheet_antennas,sheet_antEParams,sheet_sectors,sheet_secAntPorts,sheet_secAntennas,];


//  function for sheet 1
function excelltomongozero(result,it){
    console.log("connecting to mongo")
    mongoose.connect('mongodb://localhost:27017/doraemon',{ useNewUrlParser: true });
    mongoose.Promise = global.Promise;
    
     var site_tab = mongoose.model(collectlist[it], schemalist[it]);
     
     
    site_tab.insertMany(result, function(err, datas){
        if (err) throw err;
        console.log("success")
        mongoose.connection.close()
        console.log("Database update complete for :" + collectlist[it])
        console.timeEnd('Time Taken to Import '+ tabs[it]);
        readexceltojson(1);

    });
    
}


//  function for sheet 2
function excelltomongoone(result,it){
    console.log("connecting to mongo")
    mongoose.connect('mongodb://localhost:27017/doraemon',{ useNewUrlParser: true });
    mongoose.Promise = global.Promise;
    
     var site_tab = mongoose.model(collectlist[it], schemalist[it]);
     
     
    site_tab.insertMany(result, function(err, datas){
        if (err) throw err;
        console.log("success")
        mongoose.connection.close()
        console.log("Database update complete for :" + collectlist[it])
        console.timeEnd('Time Taken to Import '+ tabs[it]);
        readexceltojson(2);

    });
    
}


//  function for sheet 3
function excelltomongotwo(result,it){
    console.log("connecting to mongo")
    mongoose.connect('mongodb://localhost:27017/doraemon',{ useNewUrlParser: true });
    mongoose.Promise = global.Promise;
    
     var site_tab = mongoose.model(collectlist[it], schemalist[it]);
     
     
    site_tab.insertMany(result, function(err, datas){
        if (err) throw err;
        console.log("success")
        mongoose.connection.close()
        console.log("Database update complete for :" + collectlist[it])
        console.timeEnd('Time Taken to Import '+ tabs[it]);
        readexceltojson(3);

    });
    
}


//  function for sheet 4
function excelltomongothree(result,it){
    console.log("connecting to mongo")
    mongoose.connect('mongodb://localhost:27017/doraemon',{ useNewUrlParser: true });
    mongoose.Promise = global.Promise;
    
     var site_tab = mongoose.model(collectlist[it], schemalist[it]);
     
     
    site_tab.insertMany(result, function(err, datas){
        if (err) throw err;
        console.log("success")
        mongoose.connection.close()
        console.log("Database update complete for :" + collectlist[it])
        console.timeEnd('Time Taken to Import '+ tabs[it]);
        readexceltojson(4);

    });
    
}


//  function for sheet 5
function excelltomongofour(result,it){
    console.log("connecting to mongo")
    mongoose.connect('mongodb://localhost:27017/doraemon',{ useNewUrlParser: true });
    mongoose.Promise = global.Promise;
    console.log("Database update complete for :" + collectlist[it])
     var site_tab = mongoose.model(collectlist[it], schemalist[it]);
     
     
    site_tab.insertMany(result, function(err, datas){
        if (err) throw err;
        console.log("success")
        mongoose.connection.close()
        console.log("Database update complete for :" + collectlist[it])
        console.timeEnd('Time Taken to Import '+ tabs[it]);
        readexceltojson(5);

    });
    
}


//  function for sheet 6
function excelltomongofive(result,it){
    console.log("connecting to mongo")
    mongoose.connect('mongodb://localhost:27017/doraemon',{ useNewUrlParser: true });
    mongoose.Promise = global.Promise;
    
     var site_tab = mongoose.model(collectlist[it], schemalist[it]);
     
     
    site_tab.insertMany(result, function(err, datas){
        if (err) throw err;
        console.log("success")
        mongoose.connection.close()
        
        console.log("Database upde complete for :" + collectlist[it])
        console.log("..."+ "\n" + "..."+ "\n" +"..."+ "\n" +"..."+ "\n" +"..."+ "\n")
        console.timeEnd('Import '+ tabs[it]);
        console.log("Excel Import Completed and Successful!!!")
        console.time('Time Taken to Update Planet to Mongo');

    });
    
}
