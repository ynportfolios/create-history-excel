const functions = require("firebase-functions");
const admin = require('firebase-admin')
const path = require('path')
const os = require("os")
const xlsx = require('xlsx');
admin.initializeApp();
exports.createHistoryExcel =  functions.https.onCall(async(data, context) => {
  const skillId = parseInt(data.skillId);
  const bucket = admin.storage().bucket();
  const templatefileName = "template.xlsx";
  const fullPath = path.join(os.tmpdir(), templatefileName);
  const now = new Date()
  const stringArray = now.toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' }).split(" ", 2)
  const date = stringArray[0].split("/").join("")
  const time = stringArray[1].split(":").join("")
  const milliSeconds = now.getMilliseconds()
  const destPath = `${skillId}/${date}/${skillId}_history_${date}${time}${milliSeconds}.xlsx`
  let histories = []
  const main = (data, context) => {
    return new Promise((resolve, reject) => {
      admin.firestore().collection('histories')
      .where('skillId', '==', skillId)
      .orderBy("createDate", "desc")
      .get().then(result => {
        result.forEach((doc) => {
          const history = {
            title: doc.data().title,
            description: doc.data().description,
            createDate: doc.data().createDate,
          }
          histories.push(history)
        });
      }).then(() => {
        bucket.file(templatefileName).download({
          destination: fullPath
        })
        .catch(err => {
          functions.logger.error("Error!", err);
        })
        .then(() => {
          let cellNumber = 1;
          let book = xlsx.readFile(fullPath);
          const sheetName = book.SheetNames[0];
          const sheet = book.Sheets[sheetName];
          histories.forEach((history) => {
            sheet[`B${cellNumber + 1}`] = { t: "s", v: history.title, w: history.title };
            sheet[`B${cellNumber + 2}`] = { t: "s", v: history.description, w: history.description };
            createDate = history.createDate.toDate().toLocaleString("ja-JP").split(" ", 1)[0];
    
            sheet[`B${cellNumber + 3}`] = { t: "s", v: createDate, w: createDate}
            cellNumber = cellNumber + 4
          });
          sheet["!ref"] = `B1:B${cellNumber}`;
          book.Sheets[sheetName] = sheet;
          xlsx.writeFile(book, fullPath);
          const metadata = { "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" };
          bucket.upload(fullPath
          , {
            destination: destPath, metadata: metadata
          }).then(function () {
            resolve(0)
          });
        });
      });
    });
  };
  await main(data, context)
  return destPath
});
