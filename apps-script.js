// วาง code นี้ใน Google Apps Script แล้ว Deploy ใหม่เป็น New Deployment
// (Execute as: Me, Who has access: Anyone)

var FOLDER_NAME = 'ใบสมัครงาน';

function getOrCreateFolder(parent, name) {
  var folders = parent ? parent.getFoldersByName(name) : DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : (parent ? parent.createFolder(name) : DriveApp.createFolder(name));
}

function saveImage(base64Data, filename, folder) {
  try {
    var clean = base64Data.replace(/^data:image\/\w+;base64,/, '');
    var mimeMatch = base64Data.match(/^data:(image\/\w+);base64,/);
    var mime = mimeMatch ? mimeMatch[1] : 'image/png';
    var ext = mime.split('/')[1] || 'png';
    var blob = Utilities.newBlob(Utilities.base64Decode(clean), mime, filename + '.' + ext);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch(e) {
    return 'error: ' + e.toString();
  }
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // สร้าง header ถ้ายังไม่มี
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'วันที่ส่ง','ชื่อ','นามสกุล','อายุ','วันที่สมัคร','ตำแหน่ง','เงินเดือน','สถานะภาพ',
        'การศึกษา','สาขาวิชา','จำนวนบุตร',
        'บ้านเลขที่','หมู่','ซอย','ถนน','ตำบล','อำเภอ','จังหวัด','อาศัยกับ',
        'เลขบัตรประชาชน','มือถือ','เฟสบุ๊ค','ไอดีไลน์',
        'ชื่อบิดา','อาชีพบิดา','ชื่อมารดา','อาชีพมารดา',
        'ชื่อสามี/ภรรยา','อาชีพสามี/ภรรยา',
        'พี่น้องทั้งหมด','พี่น้องชาย','พี่น้องหญิง',
        'บุตร','บุตรชาย','บุตรหญิง',
        'เคยงานขาย','รายละเอียดงานขาย','เปลี่ยนสาขาได้','ทำงานอาทิตย์ได้',
        'งานที่1','งานที่2',
        'ติดต่อฉุกเฉิน1','โทร1','FB1',
        'ติดต่อฉุกเฉิน2','โทร2','FB2',
        'เอกสารแนบ',
        'รูปใบสมัคร',
        'สำเนาบัตรประชาชน','สำเนาทะเบียนบ้าน','วุฒิการศึกษา','รูปถ่าย'
      ]);
    }

    var docs = [];
    if (data.docId)    docs.push('สำเนาบัตรประชาชน');
    if (data.docHouse) docs.push('สำเนาทะเบียนบ้าน');
    if (data.docEdu)   docs.push('วุฒิการศึกษา');
    if (data.docPhoto) docs.push('รูปถ่าย');

    // สร้างโฟลเดอร์ย่อยสำหรับผู้สมัครแต่ละคน
    var rootFolder = getOrCreateFolder(null, FOLDER_NAME);
    var personName = (data.fname || 'ไม่ระบุ') + '_' + (data.lname || '') + '_' + new Date().getTime();
    var personFolder = getOrCreateFolder(rootFolder, personName);

    // บันทึกรูปใบสมัคร
    var formImageUrl = '';
    if (data.imageBase64) {
      formImageUrl = saveImage(data.imageBase64, 'ใบสมัคร', personFolder);
    }

    // บันทึกรูปเอกสารแนบ
    var docIdUrl    = data.docId_img    ? saveImage(data.docId_img,    'บัตรประชาชน', personFolder)  : '';
    var docHouseUrl = data.docHouse_img ? saveImage(data.docHouse_img, 'ทะเบียนบ้าน', personFolder) : '';
    var docEduUrl   = data.docEdu_img   ? saveImage(data.docEdu_img,   'วุฒิการศึกษา', personFolder) : '';
    var docPhotoUrl = data.docPhoto_img ? saveImage(data.docPhoto_img, 'รูปถ่าย', personFolder)      : '';

    sheet.appendRow([
      new Date(),
      data.fname, data.lname, data.age, data.applyDate,
      data.position, data.salary, data.status,
      data.education, data.major, data.childrenCount,
      data.houseNo, data.moo, data.soi, data.road,
      data.tambon, data.amphoe, data.province, data.living,
      data.idCard, data.phone, data.facebook, data.lineId,
      data.fatherName, data.fatherOcc,
      data.motherName, data.motherOcc,
      data.spouseName, data.spouseOcc,
      data.siblings, data.siblingsM, data.siblingsF,
      data.ownChildren, data.ownChildrenM, data.ownChildrenF,
      data.hasSales, data.salesDetail, data.canRead, data.canSunday,
      data.work1, data.work2,
      data.em1Name, data.em1Phone, data.em1Fb,
      data.em2Name, data.em2Phone, data.em2Fb,
      docs.join(', '),
      formImageUrl,
      docIdUrl, docHouseUrl, docEduUrl, docPhotoUrl
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput('Apps Script is running.')
    .setMimeType(ContentService.MimeType.TEXT);
}
