function createFolder() {
  class Student {
    constructor(record){
      [this.grade, this.id, this.schoolId, this.school, this.department, this.classroom, 
      this.numInClass, this.sei, this.mei, this.subject, this.email] = record;
      this.name = this.sei + this.mei;
    }
  }
  setVal();
  var rootFolder  = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty("rootFolderID"));
  var subject = ["物理","化学","生物","地学","数学"];
  var studentsInfo = SpreadsheetApp.getActiveSheet().getDataRange().getValues();
  var msg = [];

  // 教科フォルダ(なければ作成)
  for (var i = 0; i < subject.length; i++){
    if (!rootFolder.getFoldersByName(subject[i]).hasNext()) {
      rootFolder.createFolder(subject[i]);
    }
  }

  // 生徒用フォルダ(なければ作成)
  for (var i = 1; i < studentsInfo.length; i++){
    var student = new Student(studentsInfo[i]);
    msg.push(addStudent(student));
  }
  sendMsg(msg);

  return 0;
}
function addStudent(student) {
  setVal();
  var str = "";
  var rootFolder  = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty("rootFolderID"));
  folders = rootFolder.getFolders();

  // フォルダ一覧から科目一致するフォルダを探す
  while(folders.hasNext()){
    folder = folders.next();
    if (student.subject == folder.getName()){
      // 当該科目フォルダを検索、該当生徒いなければ新規作成
      if (folder.getFoldersByName(student.name).hasNext()){
        str = "既成：" + student.name + "(" + student.school + ", ID:" + student.id + ")=" + student.subject;
      } else {
        studentFolder = folder.createFolder(student.name);
        // 生徒自身に権限付与
        studentFolder.addEditor(student.email);
        // カルテ作成
        createCarte(studentFolder,student);
        // 経過報告ワード作成
        createDocument(studentFolder,student);
        // 資料フォルダ作成
        studentFolder.createFolder(student.name + "_研究資料");
        str = "作成：" + student.name + "(" + student.school + ", ID:" + student.id + ")→" + student.subject;
        break;
      }
    }
  }

  if (str==""){
    str = "失敗：" + student.name + "(" + student.school + ", ID:" + student.id + ")";
  }
  return str;
}

function sendMsg(msg){
  var msgStr = "【フォルダ作成結果】";
  for (var i = 0; i < msg.length; i++){
    msgStr += "\\n" + msg[i];
  }
  Browser.msgBox(msgStr);
}

function createDocument(folder,student){
  var name = student.name;
  var docsName = name + "_研究経過報告シート";
  docs = DocumentApp.create(docsName);
  var ssId = docs.getId();

  // docsをGoogleDrive内にてあつかえる様にする。
  var file = DriveApp.getFileById(ssId);

  // GoogleDriveのフォルダをIDで指定 → そこにファイルをコピーし、ルート直下のファイルを削除
  var folder = DriveApp.getFolderById(folder.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
}

function createCarte(folder, student){
  var name = student.name;
  var docsName = name + "_個人研究カルテ";
  docs = SpreadsheetApp.create(docsName);
  var ssId = docs.getId();

  // docsをGoogleDrive内にてあつかえる様にする。
  var file = DriveApp.getFileById(ssId);

  // GoogleDriveのフォルダをIDで指定 → そこにファイルをコピーし、ルート直下のファイルを削除
  var folder = DriveApp.getFolderById(folder.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  initCarte(ssId, student);
}

function initCarte(carteID, student){
  var carte = SpreadsheetApp.openById(carteID);
  const tempID = PropertiesService.getScriptProperties().getProperty("carteSheetID");
  var temprate = SpreadsheetApp.openById(tempID);
  var sheet = temprate.getSheets()[0];
  sheet.copyTo(carte);
  carte.deleteSheet(carte.getSheets()[0]);
  var carteSheet = carte.getSheets()[0];
  editCarte(carteSheet, student);
}

function editCarte(sheet,student){
  sheet.getRange(3,4).setValue(student.school);
  sheet.getRange(3,7).setValue(student.name);
  sheet.getRange(3,12).setValue(student.furigana);
  sheet.getRange(3,16).setValue(student.department + "・" + student.classroom);
  sheet.getRange(7,1).setValue(student.subject);
}