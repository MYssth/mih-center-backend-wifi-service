require("dotenv").config({ path: `.env.${process.env.NODE_ENV}` });
var config = require("./dbconfig");
const sql = require("mssql");
// Requiring the module
const reader = require("xlsx");

async function insertUser(path) {
  try {
    console.log("insertUser call try connect to server");
    let pool = await sql.connect(config);
    console.log("connect complete");

    // Reading our test file
    const file = reader.readFile(path);

    const temp = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[0]]);
    temp.forEach(async (res) => {
      // console.log(res);
      if (
        res.UWFINSDATE !== undefined &&
        res.UWFUSRNAM !== undefined &&
        res.UWFPASSWORD !== undefined
      ) {
        await pool
          .request()
          .input("UWFINSDATE", sql.VarChar, res.UWFINSDATE)
          .input("UWFUSRNAM", sql.VarChar, res.UWFUSRNAM)
          .input("UWFPASSWORD", sql.VarChar, res.UWFPASSWORD)
          .query(
            "INSERT INTO USRWIFIV5PF" +
              " (UWFINSDATE" +
              ", UWFUSRNAM" +
              ", UWFPASSWORD)" +
              " VALUES" +
              " (@UWFINSDATE" +
              ", @UWFUSRNAM" +
              ", @UWFPASSWORD)"
          );
      }
    });
    console.log("insertUser complete");
    console.log("====================");
    return { status: "ok", message: "เพิ่มรหัส Wi-Fi เรียบร้อย" };
  } catch (error) {
    console.error(error);
    return { status: "error", message: error.message };
  }
}

async function getVersion() {
  try {
    return process.env.version;
  } catch (error) {
    console.error(error);
    return { status: "error", message: error.message };
  }
}

module.exports = {
  insertUser: insertUser,
  getVersion: getVersion,
};
