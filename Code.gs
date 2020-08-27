function onEdit() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var totalMiles = [0, 0, 0, 0, 0, 0];
  var totalPace = [0, 0, 0, 0, 0, 0];
  var totalNum = [0, 0, 0, 0, 0, 0];
  var totalRest = [0, 0, 0, 0];
  var totalDays = [0, 0];
  var itemLists = [[], [], [], [], [], []]; //List order: weathers, partners, conditions, time of day, shoes, location
  var countLists = [[], [], [], [], [], []]; //Pulled from spreadsheet columns 7, 9, 10, 11, 12, 13
  var locDistances = [];
  var locTimes = [];
  var locPaces = [];
  var mergedLists = [[], [], [], [], [], []];
  var colors = ['#d9ead3', '#d9d2e9', '#fff2cc', '#ffe599', '#f9cb9c', '#fce5cd'];
  var color, i, j, k, n, s, col;
  var presets = [[], [], [], [], []]; //weather, conditions, time of day, shoes, location
  
  presets[0] = {".": "clear", "c": "cloudy", "o": "overcast", "l": "light rain", "h": "heavy rain", "i": "inside", "f": "foggy", "s": "snow"};
  presets[1] = {".": "dry", "w": "wet", "i": "inside", "c": "icy", "m": "muddy"};
  presets[2] = {".": "morning", "e": "early", "a": "afternoon", "l": "late"};
  presets[3] = {".": "Hoka Clifton 6", "t": "New Balance spikes", "x": "Asics spikes", "v": "Nike Vaporfly Next%"};
  presets[4] = {".": "centennial", "np": "nike>powerline", "nn": "nike>nature", "sb": "suburbia", "ir": "indoor rec", "or": "outdoor rec", 
                "gw": "greenway", "bm": "belle meade", "lc": "love circle", "mc": "mccabe", "sl": "shoreline", "ws": "watershed", "hm": "hartman", "rp": "rose park", 
                "eh": "english hill", "pl": "pipeline", "rh": "RHS track", "12": "12 south", "ls": "lake samm", "rc": "rollercoaster", "rt": "river trail",
                "p1": "powerline 1", "p2": "powerline 2", "p3": "powerline 3", "ev": "everest"};
  
  for (i=2;i<367;i++) { //iterates through each day
    if(data[i][7].charAt(0) == ".") { //uses code presets to populate info cells if a code was given
      sheet.getRange(i+1,8).setValue(presets[0][data[i][7].charAt(1)]);
      sheet.getRange(i+1,10).setValue("solo");
      sheet.getRange(i+1,11).setValue(presets[1][data[i][7].charAt(2)]);
      sheet.getRange(i+1,12).setValue(presets[2][data[i][7].charAt(3)]);
      sheet.getRange(i+1,13).setValue(presets[3][data[i][7].charAt(4)]);
      sheet.getRange(i+1,14).setValue(presets[4][data[i][7].slice(5)]);
    }
    
    //totals distances and paces for workouts where min/mile pace is stated
    color = sheet.getRange(i+1,9).getBackground();
    if(colors.indexOf(color) != -1) {
      n = colors.indexOf(color);
      if (n<2) { //if recovery or long run, gets distance from total workout distance, and adds to total
        totalMiles[n] += data[i][4];
        if(data[i][8].slice(data[i][8].length - 1) == "@") { //autofills pace from time and distance if needed
          sheet.getRange(i+1,9).setValue(data[i][8] + asTime(data[i][5]*60/data[i][4]));
          data[i][8] = data[i][8] + asTime(data[i][5]*60/data[i][4]); //to update for later in the script
        }
      } else { //if quality workout, gets distance from workout notes, converts to miles, and adds distance and rest time to total
        totalMiles[n] += data[i][8].slice(data[i][8].indexOf(",") + 2, data[i][8].indexOf("@") - 1)/1609;
        totalRest[n-2] += asSecs(data[i][8], data[i][8].indexOf("(") + 1);
      }
      totalPace[n] += asSecs(data[i][8], data[i][8].indexOf("@") + 1);
      totalNum[n]++; //increments the number of workouts of that type
    }
    
    //handles listing of weather, partners, conditions, time of day, and location
    for (j=0;j<6;j++) { //for clarity, j is the identifier of the type of list being made
      col = (j == 0) ? 7 : (j + 8); //function to map this identifier to the actual column as it appears in the spreadsheet
      if (data[i][col] != "") {
        if(j == 1) { //handles separating lists of partner names if necessary
          s = data[i][col].split(", ");
          for(k=0;k<s.length;k++) {
            if(itemLists[j].indexOf(s[k]) == -1) {
              itemLists[j].push(s[k]);
              countLists[j].push(0);
            }
            countLists[j][itemLists[j].indexOf(s[k])]++;
          }
        } else { //handles all other lists with only one item per cell
          if(itemLists[j].indexOf(data[i][col]) == -1) {
            itemLists[j].push(data[i][col]);
            countLists[j].push(0);
            if(j == 5) {
              locDistances.push(0);
              locTimes.push(0);
            }
          }
          countLists[j][itemLists[j].indexOf(data[i][col])]++;
          if(j == 5) {
            locDistances[itemLists[j].indexOf(data[i][col])] += data[i][4];
            locTimes[itemLists[j].indexOf(data[i][col])] += data[i][5];
          }
        }
      }
    }
  }
  
  for (i=0;i<locDistances.length;i++) { //finds pace for each location based on time and distance
    locPaces[i] = asTime(60 * locTimes[i] / locDistances[i]);
  }
  
  for (i=0;i<6;i++) { //sorts and prints each list with its count
    for(j=0;j<itemLists[i].length;j++) {
      if(i<5) { //concatenates items and counts into single objects to be sorted, and pace if applicable
        mergedLists[i].push({count: countLists[i][j], item: itemLists[i][j]});
      } else {
        mergedLists[i].push({count: countLists[i][j], item: itemLists[i][j], pace: locPaces[j]});
      } 
    }
    mergedLists[i] = mergedLists[i].sort(function(a,b){return b.count - a.count});
    for(j=0;j<mergedLists[i].length;j++) {
      sheet.getRange(18+j,16+(2*i)).setValue(mergedLists[i][j].item);
      sheet.getRange(18+j,17+(2*i)).setValue(mergedLists[i][j].count);
      if(i==5) {
        sheet.getRange(18+j,18+(2*i)).setValue(mergedLists[i][j].pace);
      }
    }
    for(j=0;j<4;j++) { //erases any list entry that could have been left behind
      sheet.getRange(18+mergedLists[i].length+j,16+(2*i)).setValue("");
      sheet.getRange(18+mergedLists[i].length+j,17+(2*i)).setValue("");
      if(i==5) {
        sheet.getRange(18+mergedLists[i].length+j,18+(2*i)).setValue("");
      }
    }
  }
  
  for (i=0;i<6;i++) {
    sheet.getRange(i+10,17).setValue(Math.round(totalMiles[i] * 10)/10); //prints total distance rounded to tenths of a mile
    sheet.getRange(i+10,18).setValue(totalNum[i] == 0 ? "" : asTime(totalPace[i]/totalNum[i])); //prints average pace, or zero if no workouts were completed of this type
    sheet.getRange(i+10,19).setValue((i > 1) ? totalNum[i] == 0 ? "" : asTime(totalRest[i-2]/totalNum[i]) : ""); //prints average rest time for quality workouts, or zero if no workouts were completed of this type
  }
}

function asSecs(s, index) { //converts a time in x:xx format to seconds
  var secs = s.slice(index+2, index+4);
  secs = (secs.charAt(0) == "0") ? parseInt(secs.slice(1, 2)) : parseInt(secs); //removes leading zero on seconds to avoid being interpreted in octal
  return parseInt(s.slice(index, index+1) * 60) + secs;
}

function asTime(secs) { //converts a time in seconds to x:xx format
  return Math.floor(secs/60) + ":" + ((secs % 60) < 10 ? "0" : "") + Math.floor(secs % 60);
}