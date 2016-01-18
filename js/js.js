var ddate, ddelay, ttime, tdelay;
function debuteDate() {
  if (!document.layers&&!document.all&&!document.getElementById)
  return
  var adate, date, amonth;
  ddelay = 10000;
  adate = new Date();
  date = adate.getDate();
  amonth = adate.getMonth()+1;
  if (amonth == 1) date += " Jan";
  else if (amonth == 2) date += " Feb";
  else if (amonth == 3) date += " Mar";
  else if (amonth == 4) date += " Apr";
  else if (amonth == 5) date += " May";
  else if (amonth == 6) date += " Jun";
  else if (amonth == 7) date += " Jul";
  else if (amonth == 8) date += " Aug";
  else if (amonth == 9) date += " Sep";
  else if (amonth == 10) date += " Oct";
  else if (amonth == 11) date += " Nov";
  else if (amonth == 12) date += " Dec";
  if (adate.getYear() > 1999)
    date += " " + adate.getYear();
  else date += "  " + (1900 + adate.getYear());
  date = "  " + date;
  if (document.layers){
  document.layers.jour.document.write(date)
  document.layers.jour.document.close()
  }
  else if (document.all)
  jour.innerHTML=date
  else if (document.getElementById)
  document.getElementById("jour").innerHTML=date
  ddate = setTimeout("debuteDate(ddelay)",ddelay);
}

function debuteTemps() {
  if (!document.layers&&!document.all&&!document.getElementById)
  return
  var hhmmss = "  ", mymin, sec;
  tdelay = 1000;
  adate = new Date();
  hhmmss += adate.getHours();
  mymin = adate.getMinutes();
  if (mymin < 10) hhmmss += ":0" + mymin;
  else hhmmss += ":" + mymin;
  sec = adate.getSeconds();
  if (sec < 10) hhmmss += ":0" + sec;
  else hhmmss += ":" + sec;
  hhmmss = " " + hhmmss;
  if (document.layers){
  document.layers.heure.document.write(hhmmss)
  document.layers.heure.document.close()
  }
  else if (document.all)
  heure.innerHTML=hhmmss
  else if (document.getElementById)
  document.getElementById("heure").innerHTML=hhmmss
  ttime = setTimeout("debuteTemps(tdelay)",tdelay);
}