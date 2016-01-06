//global variables

var subject;
var criteria;
var n_class;
var day;
var month;
var year;
var stats;
var sender;
var studentMail = "chaithrayenikapati@gmail.com";
students_status = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0];
function initialize_students()
{
  for(var i=0;i<students_status.length;i++)
  {
    students_status[i] = 0;
    Logger.log(students_status[i]);
  }
}

// sends an email to the student mail id with student statistics spreadsheet as attachment
function mailSheetToStudents() {
 
  var files = DriveApp.getFilesByName("STUDENT_STATISTICS");
  while(files.hasNext())
  {
    var file = files.next();
    MailApp.sendEmail(studentMail, "sheet", "stats",{
     name: 'Attendance update',
     attachments: [file]
 });
    Logger.log("mail sent");
  } 
}

function create_stats_sheet_b()
{
   var sheet = SpreadsheetApp.create("STATISTICS",200,5);
   stats = sheet.getId();
   Logger.log("Spreadsheet id ="+ stats);
   Logger.log("spreadsheet name="+sheet.getName());
   sheet =  SpreadsheetApp.openById(stats).getActiveSheet();
   var reg_start=61;
   var reg_end=99;
   var le_start=13;
   var le_end=24;
   sheet.appendRow(["ATTENDANCE STATISTICS"]);
   sheet.appendRow(["HTNO","NUM_OF_CLASSES","ATTENDANCE %"]);
   sheet.appendRow(["TOTAL"]);
   for(i=reg_start; i <= 99;i++)
    {
      sheet.appendRow([i]);
    }
    for(i=0;i<=9;i++)
    {
      sheet.appendRow(["A"+i]);
    } 
    for(i=0;i<=9;i++)
    {
      sheet.appendRow(["B"+i]);
    } 
    sheet.appendRow(["C"+0]);
    for(i=le_start;i<=le_end;i++)
    {
      sheet.appendRow(["le_"+i]);
    }
  
}


function create_stats_sheet_a()
{
   var sheet = SpreadsheetApp.create("STATISTICS",200,5);
   stats = sheet.getId();
   Logger.log("Spreadsheet id ="+ stats);
   Logger.log("spreadsheet name="+sheet.getName());
   sheet =  SpreadsheetApp.openById(stats).getActiveSheet();
   var reg_start=1;
   var reg_end=60;
   var le_start=1;
   var le_end=22;
   sheet.appendRow(["ATTENDANCE STATISTICS"]);
   sheet.appendRow(["HTNO","NUM_OF_CLASSES","ATTENDANCE %"]);
   sheet.appendRow(["TOTAL"]);
   for(i=reg_start; i <= reg_end;i++)
   {
      sheet.appendRow([i]);
   }
    
   for(i=le_start;i<=le_end;i++)
   {
      sheet.appendRow(["le_"+i]);
   }
   
  
}
 

function create_student_stats_sheet_b()
{
   var sheet = SpreadsheetApp.create("STUDENT_STATISTICS",200,10);
   stats = sheet.getId();
   Logger.log("Spreadsheet id ="+ stats);
   Logger.log("spreadsheet name="+sheet.getName());
   sheet =  SpreadsheetApp.openById(stats).getActiveSheet();
   var reg_start=61;
   var reg_end=99;
   var le_start=13;
   var le_end=24;
   subjects = ["","SAN","MS","SL"];
   sheet.appendRow(["ATTENDANCE STATISTICS"]);
   sheet.appendRow(subjects);
   sheet.appendRow(["TOTAL CLASSES"]);
   for(i=reg_start; i <= 99;i++)
    {
      sheet.appendRow([i]);
    }
    for(i=0;i<=9;i++)
    {
      sheet.appendRow(["A"+i]);
    } 
    for(i=0;i<=9;i++)
    {
      sheet.appendRow(["B"+i]);
    } 
    sheet.appendRow(["C"+0]);
    for(i=le_start;i<=le_end;i++)
    {
      sheet.appendRow(["le_"+i]);
    }
  
}


function create_student_stats_sheet_a()
{
   var sheet = SpreadsheetApp.create("STUDENT_STATISTICS",200,10);
   stats = sheet.getId();
   Logger.log("Spreadsheet id ="+ stats);
   Logger.log("spreadsheet name="+sheet.getName());
   sheet =  SpreadsheetApp.openById(stats).getActiveSheet();
   var reg_start=1;
   var reg_end=60;
   var le_start=1;
   var le_end=22;
  
   subjects = ["","SAN","MS","SL"];
   sheet.appendRow(["ATTENDANCE STATISTICS"]);
   sheet.appendRow(subjects);
   sheet.appendRow(["TOTAL CLASSES"]);
   for(i=reg_start; i <= reg_end;i++)
   {
      sheet.appendRow([i]);
   }
    
   for(i=le_start;i<=le_end;i++)
   {
      sheet.appendRow(["le_"+i]);
   }
   
  
}

// extracts email id of the sender from the sender info
function c_function1(full_sender)
{
  var email_id= '';
  var flag=0;
    for(var j=0; full_sender[j] != '>'; j++)
    {
       if( full_sender[j] == '<')
       {
          flag=1;
          continue;
       }
      else if(flag == 1)
        email_id+=full_sender[j];
    }
    sender = email_id;
    return email_id;
}

// opens the desired spreadsheet and returns to the calling function
function sheets()
{
   var id = "";
   var files = DriveApp.getFiles();
   while (files.hasNext()) 
   {
     var file = files.next();
     if (file.getName() == subject)
     {
       id = file.getId();
       var sheet = SpreadsheetApp.openById(id).getActiveSheet();
       return sheet;
     }
   }
  return -1;
}

function intializeDate()
{
  var subjects=["SAN","SL","MS"];
  for( i=0; i<subjects.length; i++)
  {
    subject=subjects[i];
    var sheet = sheets();
    for(j=1; j<=31; j++)
    {
      var data = sheet.getDataRange().getValues();
      sheet.getRange(2,j+1).setValue("3/"+j+"/2014");
    
    }
  }
}

//var id;
function create_sample_spreadsheet()
{
   subjects=["SAN","SL","MS"];
   var i;
   for(i=0 ; i<subjects.length ; i++)
   {
     
     var sheet = SpreadsheetApp.create(subjects[i],200,20);
     var id = sheet.getId();
     Logger.log("Spreadsheet id ="+ id);
     
     Logger.log("spreadsheet name="+sheet.getName());
   }

}
 

/*var reg_start=61;
var reg_end=120;
var le_start=13;
var le_end=24;*/
function initialise_spreadsheet_section_A()
{
  var subjects=["SAN","SL","MS"];
  for( j=0; j<subjects.length; j++)
  {
    subject=subjects[j];
    var sheet = sheets();
    var i;
    var reg_start=1;
    var reg_end=60;
    var le_start=1;
    var le_end=12;
    sheet.appendRow([subject]);
    sheet.appendRow(["DATE"]);
    sheet.appendRow(["TOT_CLASSES"]);
    for(i=reg_start; i <= reg_end;i++)
    {
      sheet.appendRow([i]);
    }
    
    for(i=le_start;i<=le_end;i++)
    {
      sheet.appendRow(["le_"+i]);
    }
  }
}

function initialise_spreadsheet_section_B()
{
  var subjects=["SAN","SL","MS"];
  for( j=0; j<subjects.length; j++)
  {
    subject=subjects[j];
    var sheet = sheets();
    var i;
    var reg_start=61;
    var reg_end=120;
    var le_start=13;
    var le_end=24;
    sheet.appendRow([subject]);
    sheet.appendRow(["DATE"]);
    sheet.appendRow(["TOT_CLASSES"]);
    for(i=reg_start; i <= 99;i++)
    {
      sheet.appendRow([i]);
    }
    for(i=0;i<=9;i++)
    {
      sheet.appendRow(["A"+i]);
    } 
    for(i=0;i<=9;i++)
    {
      sheet.appendRow(["B"+i]);
    } 
    sheet.appendRow(["C"+0]);
    for(i=le_start;i<=le_end;i++)
    {
      sheet.appendRow(["le_"+i]);
    }
  }
}

function initialiseTotalClasses()
{
  var subjects=["SAN","SL","MS"];
  for( j=0; j<subjects.length; j++)
  {
    subject=subjects[j];
    var sheet = sheets();
    
    for(i = 1; i<= 31; i++)
    {
     var data = sheet.getDataRange().getValues();
     sheet.getRange(3,i+1).setValue(0);
    }
  }
}

function update_total_classes()
{
  var nth_class = 6;
  var col = 2;
  sheet = SpreadsheetApp.openById("tdtawjnukYcGsFPruukM6Uw").getActiveSheet();
  var data = sheet.getDataRange().getValues();
  sheet.getRange(3,col).setValue(nth_class);
}

function getColumn()
{
  Logger.log("getColumn");
  //sheet = SpreadsheetApp.openById("tdtawjnukYcGsFPruukM6Uw").getActiveSheet();
  var sheet = sheets(); 
  var data = sheet.getDataRange().getValues();
  for(i = 1;i < data[1].length;i++)
  {
     var m = data[1][i].getMonth()+1;
     var y = data[1][i].getYear();
     dt = data[1][i].toString();
     var d = "";
     if(dt[8]!= "0")
     {
       d += dt[8];
     }
     d += dt[9];
     if(year==y && month==m && day==d)
     { 
       return i+1;
     }
  }
  return -1;
}

function integrate()
{
  var inbox_threads = GmailApp.getInboxThreads();
  Logger.log("integrate");
  Logger.log("no of threads:"+inbox_threads.length);
  for(var i=0;i<inbox_threads.length;i++)
  {
     Logger.log("for loop");
     Logger.log(i);
     var first_message = inbox_threads[i].getMessages()[0];
     c_function1(first_message.getFrom());
     if(isLecturer(sender) && first_message.isUnread())
     {   
       Logger.log(first_message.getPlainBody());
       first_message.markRead();
       Logger.log("marked read");
       var mailBody = first_message.getPlainBody();
       var array = parseMailBody( mailBody );
       Logger.log("subject: "+subject);
       var result;
       if(array == "no of classes limit exceeded" || array == "invalid date format" || array == "invalid mail format")
           MailApp.sendEmail(sender, "error_ack", array);
       else
       {
         if(criteria == "a" || criteria == "undo_a")
         {
           Logger.log(criteria);
           if(criteria == "undo_a")
             n_class = -n_class;
           result = update_absent(array);
           if(result == "success")
           {
             update_stud_stats_a(array);
             Logger.log("absenties updation");
             MailApp.sendEmail(sender, "acknowledgement", "attendance updated: "+criteria);
           }
           else
           {
              MailApp.sendEmail(sender, "error_ack", "update failed");
           }
         }
         if(criteria == "p" || criteria == "undo_p")
         {
           Logger.log(criteria);
           if(criteria == "undo_p")
             n_class = -n_class;
           result = update_present(array);
           if(result == "success")
           {
             update_stud_stats_p(array);
             MailApp.sendEmail(sender, "acknowledgement", "attendance updated: "+criteria);
           }
           else
           {
              MailApp.sendEmail(sender, "error_ack", "update failed");
           }
          }
          if(criteria == "c+" || criteria == "undo_c-")
          {
            Logger.log("correction");
            result = correction_add_attendance(array);
            if(result == "success")
            {
              update_stud_stats_cadd(array);
              MailApp.sendEmail(sender, "ack", "attendance corrected: "+criteria);
            }
            else
            {
              MailApp.sendEmail(sender, "error_ack", "update failed");
            }
          }
          if(criteria == "c-" || criteria == "undo_c+")
          {
            Logger.log("correction");
            result = correction_subtract_attendance(array);
            if(result == "success")
            {
              update_stud_stats_csub(array);
              MailApp.sendEmail(sender, "ack", "attendance corrected: "+criteria);
            }
            else
            {
              MailApp.sendEmail(sender, "error_ack", "update failed"); 
            }
          }
       }
       
     }
  } 
  update_stats_percent();
  Logger.log("end of execution");
}

function update_present(presentees)
{
  initialize_students();
  Logger.log("update_present");
  //sheet = SpreadsheetApp.openById("tdtawjnukYcGsFPruukM6Uw").getActiveSheet();
  Logger.log(subject);
  var sheet = sheets(); 
  
  var sub_temp = subject;
  
  //open stats sheet
  subject = "STATISTICS";
   Logger.log(subject);
  var stats = sheets();
  Logger.log(stats);
  //var buf = stats.getDataRange().getValues();
  
  subject = sub_temp;
   Logger.log(subject);
  if(sheet == -1)
    return "invalid subject name";
  var data = sheet.getDataRange().getValues();
  var col= getColumn();
  if(col==-1)
    return "failure";
  Logger.log("col:"+col);
  
  
  // add n_class value to total classes field of statistics sheet
  temp = stats.getRange(3,2).getValue();
  if(temp=="")
  {
       temp=0;
  }
  else
  {
      temp=parseFloat(temp);
  }
  temp+=n_class;
  temp = parseFloat(temp);
  stats.getRange(3,2).setValue(temp); 
  
  
  
  var total_classes_taken = sheet.getRange(3,col).getValue();
  var num_of_classes=n_class;
  var row=4;
  total_classes_taken += n_class;
  sheet.getRange(3,col).setValue(total_classes_taken);  
  for (var i=row; i<=data.length;i++){
    
    for(var j=0; j<presentees.length; j++){
      if(data[i-1][0]==presentees[j].trim() && students_status[i-4] == 0)
      {
        students_status[i-4] = 1;
        //Logger.log(data[2][3]);
        Logger.log(data[i-1][0]);
        var temp=sheet.getRange(i,col).getValue();
        if(temp=="")
        {
          sheet.getRange(i,col).setValue(num_of_classes);
        }
        else
        {
          temp=parseFloat(temp);
        
          temp=temp+num_of_classes;
        //Logger.log("\n",parseInt(temp));
          sheet.getRange(i,col).setValue(temp);
          
        }
        
        //update stats sheet for this roll num
        
        var temp=stats.getRange(i,2).getValue();
        if(temp=="")
        {
          stats.getRange(i,2).setValue(num_of_classes);
        }
        else
        {
          temp=parseFloat(temp);
        
          temp=temp+num_of_classes;
        //Logger.log("\n",parseInt(temp));
          stats.getRange(i,2).setValue(temp);
          
        }
                

   
      }
    }
    
  }
  return "success";
}

function update_absent(absentees)
{
  //var absentees = [64,73,85,96,106];
  //var n_class=6;
  //sheet = SpreadsheetApp.openById("tdtawjnukYcGsFPruukM6Uw").getActiveSheet();
  initialize_students();
  var sheet = sheets();
  
  var sub_temp = subject;
  
  //open stats sheet
  subject = "STATISTICS";
  Logger.log(subject);
  var stats = sheets();
  Logger.log(stats);
  //var buf = stats.getDataRange().getValues();
  
  subject = sub_temp;
   Logger.log(subject);
  
  
  if(sheet == -1)
    return "invalid subject name";
  var data = sheet.getDataRange().getValues();
  var col=getColumn();
  Logger.log(col);
  
 
  
  
  if(col==-1)
    return "failure";
  
  
   // add n_class value to total classes field of statistics sheet
  temp = stats.getRange(3,2).getValue();
  if(temp=="")
  {
       temp=0;
  }
  else
  {
      temp=parseFloat(temp);
  }
  temp+=n_class;
  temp = parseFloat(temp);
  stats.getRange(3,2).setValue(temp); 
  
  
  
  var total_classes_taken=sheet.getRange(3,col).getValue();
  var num_of_classes=n_class;
  var row=4;
  total_classes_taken += n_class;
  sheet.getRange(3,col).setValue(total_classes_taken); 
  var flag=true;
  
  for (var i=row; i<=data.length;i++)
  {  
    flag=true;
    for(var j=0; j<absentees.length; j++)
    {
      if(data[i-1][0]==absentees[j].trim() && students_status[i-4] == 0)
      {
        students_status[i-4] = 1;
        flag=false;
        break;
      }
    }
    if(flag)
    {
        var temp=sheet.getRange(i,col).getValue();
      
        if(temp=="")
        {
          sheet.getRange(i,col).setValue(num_of_classes);
        }
        else
        {
          temp=parseFloat(temp);
          temp=temp+num_of_classes;
        //Logger.log("\n",parseInt(temp));
          sheet.getRange(i,col).setValue(temp);
        }
      
        //update stats sheet for this roll num
        
        var temp=stats.getRange(i,2).getValue();
        if(temp=="")
        {
          stats.getRange(i,2).setValue(num_of_classes);
        }
        else
        {
          temp=parseFloat(temp);
        
          temp=temp+num_of_classes;
        //Logger.log("\n",parseInt(temp));
          stats.getRange(i,2).setValue(temp);
          
        }
    } 
  } 
 return "success";
}

function correction_add_attendance(students)
{
  // function to correct already entered attendance
  // for adding attendance: num_of_classes is +ve value
  // for subtracting attendance: num_of_classes is -ve value
  // number of classes present = current attendance + num_of_classes
  //var students = [99];
  initialize_students();
  var sheet = sheets();
  
  
  if(sheet == -1)
    return "invalid subject name";
  
  var sub_temp = subject;
  
  //open stats sheet
  subject = "STATISTICS";
  Logger.log(subject);
  var stats = sheets();
  Logger.log(stats);
  //var buf = stats.getDataRange().getValues();
  
  subject = sub_temp;
   Logger.log(subject);
  
  var num_of_classes = n_class;
  var row = 4;
  var col = getColumn();
  if(col==-1)
    return "failure";
  //sheet = SpreadsheetApp.openById("tdtawjnukYcGsFPruukM6Uw").getActiveSheet();
  var total_classes_taken=sheet.getRange(3,col).getValue();
  if(num_of_classes <= total_classes_taken){
  var data = sheet.getDataRange().getValues();
  for (var i=row; i<=data.length;i++){
    
    for(var j=0; j<students.length; j++){
      if(data[i-1][0].toString() == students[j].trim() && students_status[i-4] == 0 )
      {
        students_status[i-4]=1;
        Logger.log(data[i-1][0]);
        var temp=sheet.getRange(i,col).getValue();
        if(temp=="")
        {
          sheet.getRange(i,col).setValue(num_of_classes);
        }
        else
        {
          temp=parseFloat(temp);
          temp=temp+num_of_classes;
        //Logger.log("\n",parseInt(temp));
          sheet.getRange(i,col).setValue(temp);
        }
        
        
        //update stats sheet for this roll num
        
        var temp=stats.getRange(i,2).getValue();
        if(temp=="")
        {
          stats.getRange(i,2).setValue(num_of_classes);
        }
        else
        {
          temp=parseFloat(temp);
        
          temp=temp+num_of_classes;
        //Logger.log("\n",parseInt(temp));
          stats.getRange(i,2).setValue(temp);
          
        }
        
      }
    }
    
  }
    return "success";
  }
  return "failure";
}

function correction_subtract_attendance(students)
{
  // function to correct already entered attendance
  // for adding attendance: num_of_classes is +ve value
  // for subtracting attendance: num_of_classes is -ve value
  // number of classes present = current attendance + num_of_classes
  //var students = [99];
  initialize_students();
  var num_of_classes = n_class;
  var row = 4;
  var sheet = sheets();
  if(sheet == -1)
    return "invalid subject name";
  
  var sub_temp = subject;
  
  //open stats sheet
  subject = "STATISTICS";
  Logger.log(subject);
  var stats = sheets();
  Logger.log(stats);
  //var buf = stats.getDataRange().getValues();
  
  subject = sub_temp;
   Logger.log(subject);
  
  
  var col = getColumn();
  if(col==-1)
    return "failure";
  //sheet = SpreadsheetApp.openById("tdtawjnukYcGsFPruukM6Uw").getActiveSheet();
  
  var total_classes_taken=sheet.getRange(3,col).getValue();
  if(num_of_classes <= total_classes_taken){
  var data = sheet.getDataRange().getValues();
  for (var i=row; i<=data.length;i++){
    
    for(var j=0; j<students.length; j++){
      if(data[i-1][0].toString() == students[j].trim() && students_status[i-4] == 0)
      {
        students_status[i-4]=1;
        Logger.log(data[i-1][0]);
        var temp=sheet.getRange(i,col).getValue();
        if(temp=="")
        {
          //sheet.getRange(i,col).setValue(num_of_classes);
        }
        else
        {
          temp=parseFloat(temp);
          temp=temp-num_of_classes;
        //Logger.log("\n",parseInt(temp));
          sheet.getRange(i,col).setValue(temp);
        }
        
        //update stats sheet for this roll num
        
        var temp=stats.getRange(i,2).getValue();
        if(temp=="")
        {
          //stats.getRange(i,2).setValue(num_of_classes);
        }
        else
        {
          temp=parseFloat(temp);
        
          temp=temp-num_of_classes;
        //Logger.log("\n",parseInt(temp));
          stats.getRange(i,2).setValue(temp);
          
        }
        
      }
    }
    
  }
    return "success";
  }
  return "failure";
}

function update_stud_stats_p(presentees)
{
  initialize_students();
  var row = 4;
  var col;
  var temp_subject = subject;
  subject = "STUDENT_STATISTICS";
  var sheet = sheets();
  var data = sheet.getDataRange().getValues();
  for(i = 1; i< data[1].length; i++)
  {
    if( temp_subject == data[1][i]){
      col = i+1;
      break;
    }
  }
   var total_classes_taken = sheet.getRange(3,col).getValue();
  var num_of_classes=n_class;
  var row=4;
  total_classes_taken += n_class;
  sheet.getRange(3,col).setValue(total_classes_taken);  
  for (var i=row; i<=data.length;i++){
    
    for(var j=0; j<presentees.length; j++){
      if(data[i-1][0]==presentees[j].trim() && students_status[i-4] == 0)
      {
        students_status[i-4] = 1;
        //Logger.log(data[2][3]);
        Logger.log(data[i-1][0]);
        var temp=sheet.getRange(i,col).getValue();
        if(temp=="")
        {
          sheet.getRange(i,col).setValue(num_of_classes);
        }
        else
        {
          temp=parseFloat(temp);
        
          temp=temp+num_of_classes;
          sheet.getRange(i,col).setValue(temp);
        }
      }
    }
  }
}

function update_stud_stats_a(absentees)
{
  initialize_students();
  var row = 4;
  var col;
  var temp_subject = subject;
  subject = "STUDENT_STATISTICS";
  var sheet = sheets();
  var data = sheet.getDataRange().getValues();
  for(i = 1; i< data[1].length; i++)
  {
    if( temp_subject == data[1][i]){
      col = i+1;
      break;
    }
  }
  var total_classes_taken=sheet.getRange(3,col).getValue();
  var num_of_classes=n_class;
  var row=4;
  total_classes_taken += n_class;
  sheet.getRange(3,col).setValue(total_classes_taken); 
  var flag=true;
  for (var i=row; i<=data.length;i++)
  {  
    flag=true;
    for(var j=0; j<absentees.length; j++)
    {
      if(data[i-1][0]==absentees[j].trim() && students_status[i-4] == 0)
      {
        students_status[i-4] = 1;
        flag=false;
        break;
      }
    }
    if(flag)
    {
        var temp=sheet.getRange(i,col).getValue();
      
        if(temp=="")
        {
          sheet.getRange(i,col).setValue(num_of_classes);
        }
        else
        {
          temp=parseFloat(temp);
          temp=temp+num_of_classes;
        //Logger.log("\n",parseInt(temp));
          sheet.getRange(i,col).setValue(temp);
        }
    } 
  } 
}

function update_stud_stats_cadd(students)
{
  initialize_students();
  var num_of_classes = n_class;
  var row = 4;
  var col;
  var temp_subject = subject;
  subject = "STUDENT_STATISTICS";
  var sheet = sheets();
  var data = sheet.getDataRange().getValues();
  for(i = 1; i< data[1].length; i++)
  {
    if( temp_subject == data[1][i]){
      col = i+1;
      break;
    }
  }
  var total_classes_taken=sheet.getRange(3,col).getValue();
  if(num_of_classes <= total_classes_taken){
  var data = sheet.getDataRange().getValues();
  for (var i=row; i<=data.length;i++){
    
    for(var j=0; j<students.length; j++){
      if(data[i-1][0].toString() == students[j].trim() && students_status[i-4] == 0 )
      {
        students_status[i-4]=1;
        Logger.log(data[i-1][0]);
        var temp=sheet.getRange(i,col).getValue();
        if(temp=="")
        {
          sheet.getRange(i,col).setValue(num_of_classes);
        }
        else
        {
          temp=parseFloat(temp);
          temp=temp+num_of_classes;
        //Logger.log("\n",parseInt(temp));
          sheet.getRange(i,col).setValue(temp);
        }
        
        
        //update stats sheet for this roll num
        
        var temp=stats.getRange(i,2).getValue();
        if(temp=="")
        {
          stats.getRange(i,2).setValue(num_of_classes);
        }
        else
        {
          temp=parseFloat(temp);
        
          temp=temp+num_of_classes;
        //Logger.log("\n",parseInt(temp));
          stats.getRange(i,2).setValue(temp);
          
        }
        
      }
    }
    
  }
    return "success";
  }
}

function update_stud_stats_csub(students)
{
  initialize_students();
  var num_of_classes = n_class;
  var row = 4;
  var col;
  var temp_subject = subject;
  subject = "STUDENT_STATISTICS";
  var sheet = sheets();
  var data = sheet.getDataRange().getValues();
  for(i = 1; i< data[1].length; i++)
  {
    if( temp_subject == data[1][i]){
      col = i+1;
      break;
    }
  }
  var total_classes_taken=sheet.getRange(3,col).getValue();
  if(num_of_classes <= total_classes_taken){
  var data = sheet.getDataRange().getValues();
  for (var i=row; i<=data.length;i++){
    
    for(var j=0; j<students.length; j++){
      if(data[i-1][0].toString() == students[j].trim() && students_status[i-4] == 0)
      {
        students_status[i-4]=1;
        Logger.log(data[i-1][0]);
        var temp=sheet.getRange(i,col).getValue();
        if(temp=="")
        {
          //sheet.getRange(i,col).setValue(num_of_classes);
        }
        else
        {
          temp=parseFloat(temp);
          temp=temp-num_of_classes;
        //Logger.log("\n",parseInt(temp));
          sheet.getRange(i,col).setValue(temp);
        }
      }
    }
  }
    return "success";
  }
}

function update_stats_percent()
{
  
  subject = "STATISTICS";
  var stats = sheets();
  Logger.log(stats);
  var data = stats.getDataRange().getValues();
  var total = stats.getRange(3,2).getValue();
  Logger.log(total);
  for(j=4;j<=data.length ; j++)
    {
      
          
          
          var old=stats.getRange(j,2).getValue();
          if(old == "")
          {
            old = 0;
          
          }
          else
          {
            old = parseFloat(old);
          }
        
          var percent = parseFloat(old/total)*100;
          percent = parseFloat(percent);
          stats.getRange(j,3).setValue(percent);
          if(percent<= 40)
          {
            stats.getRange(j,3).setFontColor("red");
          }
          

   
    }
  return;
}

//analyzes and parses the mail body of the email and returns the array of roll numbers 
function parseMailBody(mailBody)
{
  Logger.log("parsemailbody");
   var array = mailBody.split(',');
   if(array.length <3)
     return "invalid mail format";
   subject = array[0].trim();
   criteria = array[1].trim();
   n_class = parseFloat(array[2].trim());
   if(n_class > 7)
     return "no of classes limit exceeded";
   var date = array[3].trim();
   if(date!= "")
   {
     date = date.split("/");
     if(date.length <3)
       return "invalid date format";
     month = date[0];
     day = date[1];
     year = date[2];
   }
   else
   {
     var today = new Date();
     month = today.getMonth()+1;
     year = today.getYear();
     dt = today.toString();
     day = "";
     if(dt[8]!= "0")
     {
       day += dt[8];
     }
     day += dt[9];
     Logger.log("day:"+day);
   }
   Logger.log("date updated");
   for(i=4 ;i < array.length; i++)
   {
     array[i].trim();
   }
   Logger.log(array.length);
   if((day<1|| day>31) && (month <1 || month>12) && (year <2000))
          return "invalid date format";
   return array.slice(4,array.length);
}

//testing function
function crap()
{
  var inbox_threads = GmailApp.getInboxThreads();
  var unread_count = GmailApp.getInboxUnreadCount();
  sheet = SpreadsheetApp.openById("tdtawjnukYcGsFPruukM6Uw").getActiveSheet();
  //sheet.getRange(4,2).setValue(1);
  var data = sheet.getDataRange().getValues();
  Logger.log(data[1].length);
  Logger.log(data[1][8].getMonth()+1);
  date = "3/8/2014";
  date=date.split("/");
 // Logger.log(date);
  Logger.log(data[1][3].getDay());
  if(date[0] == data[1][8].getMonth()+1 && date[2] == data[1][8].getYear())
  {
    {
      dt = data[1][8].toString();
      d="";
      if(dt[8]!= "0")
      {
        d+=dt[8];
      }
      d+=dt[9];
    }
    Logger.log(d);
    if(d==date[1])
      Logger.log("yes");
    }
   var tdy = new Date();
}

//authenticates the sender of the email
function isLecturer(sender)
{
  var lecturers = ["msrivani92@gmail.com","chaithrayenikapati@gmail.com","swathi.a278@gmail.com","shilpa.avvaru@gmail.com"];
  for(var index=0; index< lecturers.length; index++)
  {
    if(sender==lecturers[index])
      return true;
  }
  return false;
}
