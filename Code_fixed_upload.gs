const CONFIG = {

  SPREADSHEET_ID: "1cyF7nLTTzB62NAT_XzKuoO7ycP56Czqy6fMDdCAop5Y",

  DRIVE_FOLDER_ID: "1CXbplYFjxdjH1amUREFnYaIUwxeXVh9x",          // สำหรับเก็บรูปภาพ

  DRIVE_VIDEO_FOLDER_ID: "1ZCxdtKMLxob3sRHt5xU2nhnVEAgN_Ohb",    // สำหรับเก็บวิดีโอ

  ADMIN_PASSWORD: "2588",

  TIMEZONE: "Asia/Bangkok",

  SHEET_PORTFOLIO: "Portfolio",

  SHEET_SCHEDULE: "Schedule",

  SHEET_PROFILE: "Profile",

  SHEET_CALENDAR: "Calendar"

};



function driveIdFrom(url) {

  url = String(url || "");

  if (!/drive\\.google\\.com|googleusercontent\\.com/.test(url)) return url;

  var m = url.match(/[-\w]{25,}/);

  return m ? m[0] : url;

}



function toDisplayUrl(url, fileId, mediaType) {

  var id = fileId || driveIdFrom(url);

  if (!id) return String(url || "");

  if (mediaType === "video") {

    if (id.indexOf("http") === 0) return id;

    return "https\://drive.google.com/file/d/" + id + "/preview";

  }

  return "https\://lh3.googleusercontent.com/d/" + id;

}



function toImageUrl(url) {

  var id = driveIdFrom(url);

  if (!id) return String(url || "");

  if (id.indexOf("http") === 0) return id;

  return "https\://lh3.googleusercontent.com/d/" + id;

}



function toImageUrlList(csv) {

  return String(csv || "")

    .split(",")

    .map(function (s) { return s.trim(); })

    .filter(function (s) { return Boolean(s); })

    .map(toImageUrl)

    .join(",");

}



function responseJSON(data) {

  return ContentService.createTextOutput(JSON.stringify(data))

    .setMimeType(ContentService.MimeType.JSON);

}



function doOptions(e) {

  return responseJSON({ success: true, message: "CORS OK" });

}



function doGet(e) {

  try {

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);



    let profSheet = ss.getSheetByName(CONFIG.SHEET_PROFILE);

    if (!profSheet) {

      profSheet = ss.insertSheet(CONFIG.SHEET_PROFILE);

      profSheet.appendRow(["Key", "Value"]);

      const defaults = [

        ["brandName", "PHOTOGRAPHER"],

        ["heroTitle", "Luxury Visual Experience"],

        ["heroSubtitle", "รับสร้างสรรค์ผลงานภาพถ่ายระดับพรีเมียม และงานผลิตวิดีโอ Cinematic ครบวงจร"],

        ["avatarUrl", ""],

        ["featuredPhotos", ""],

        ["photographerName", "PRO PHOTOGRAPHER"],

        ["bio", "ช่างภาพมืออาชีพ & Filmmaker ประสบการณ์กว่า 8 ปี เน้นงานคุณภาพพรีเมียม สื่ออารมณ์และเรื่องราวอย่างมีสไตล์"],

        ["location", "กรุงเทพฯ & รับงานทั่วประเทศ"],

        ["phone", ""],

        ["line", "@yourlineid"],

        ["instagram", "instagram.com/yourhandle"],

        ["facebook", "facebook.com/yourpage"],

        ["tiktok", "tiktok.com/@yourprofile"],

        ["email", ""],

        ["socialLinks", "[]"]

      ];

      defaults.forEach(r => profSheet.appendRow(r));

    }

    const profRows = profSheet.getDataRange().getValues();

    const profile = {};

    for (let i = 1; i < profRows.length; i++) {

      if (profRows[i][0]) profile[profRows[i][0]] = profRows[i][1];

    }

    profile.avatarUrl = toImageUrl(profile.avatarUrl);

    profile.featuredPhotos = toImageUrlList(profile.featuredPhotos);



    let pSheet = ss.getSheetByName(CONFIG.SHEET_PORTFOLIO);

    if (!pSheet) {

      pSheet = ss.insertSheet(CONFIG.SHEET_PORTFOLIO);

      pSheet.appendRow(["ID", "Title", "Category", "MediaType", "FileUrl", "FileId", "ExtLink", "CreatedAt", "Details", "DisplayOrder"]);

    }

    const pRows = pSheet.getDataRange().getValues();

    const portfolio = [];

    for (let i = 1; i < pRows.length; i++) {

      if (pRows[i][0]) {

        const mType = pRows[i][3] || "image";

        const fId = pRows[i][5] || driveIdFrom(pRows[i][4]);

        portfolio.push({

          id: pRows[i][0],

          title: pRows[i][1],

          category: pRows[i][2],

          mediaType: mType,

          fileUrl: toDisplayUrl(pRows[i][4], fId, mType),

          thumbUrl: fId ? ("https\://lh3.googleusercontent.com/d/" + fId) : String(pRows[i][4] || ""),

          fileId: fId,

          extLink: pRows[i][6] || "",

          createdAt: pRows[i][7],

          details: pRows[i][8] || "",

          displayOrder: pRows[i][9] !== "" && pRows[i][9] !== undefined ? Number(pRows[i][9]) : 9999

        });

      }

    }

    portfolio.sort((a, b) => a.displayOrder - b.displayOrder);



    let sSheet = ss.getSheetByName(CONFIG.SHEET_SCHEDULE);

    if (!sSheet) {

      sSheet = ss.insertSheet(CONFIG.SHEET_SCHEDULE);

      sSheet.appendRow(["ID", "Date", "Time", "Title", "Location", "Status", "ExtLink", "Details", "CreatedAt"]);

    }

    const sRows = sSheet.getDataRange().getValues();

    const schedule = [];

    for (let i = 1; i < sRows.length; i++) {

      if (sRows[i][0]) {

        let rawDate = sRows[i][1];

        let formattedDate = rawDate instanceof Date

          ? Utilities.formatDate(rawDate, CONFIG.TIMEZONE, "yyyy-MM-dd")

          : String(rawDate).replace(/^'/, '');

        schedule.push({

          id: sRows[i][0],

          date: formattedDate,

          time: sRows[i][2],

          title: sRows[i][3],

          location: sRows[i][4],

          status: sRows[i][5],

          extLink: sRows[i][6] || "",

          details: sRows[i][7] || "",

          createdAt: sRows[i][8]

        });

      }

    }



    let cSheet = ss.getSheetByName(CONFIG.SHEET_CALENDAR);

    if (!cSheet) {

      cSheet = ss.insertSheet(CONFIG.SHEET_CALENDAR);

      cSheet.appendRow(["Date", "StatusType", "AvailableSlots", "Note"]);

    }

    const cRows = cSheet.getDataRange().getValues();

    const calendar = [];

    for (let i = 1; i < cRows.length; i++) {

      if (cRows[i][0]) {

        let rawDate = cRows[i][0];

        let formattedDate = rawDate instanceof Date

          ? Utilities.formatDate(rawDate, CONFIG.TIMEZONE, "yyyy-MM-dd")

          : String(rawDate).replace(/^'/, '');

        calendar.push({

          date: formattedDate,

          statusType: cRows[i][1],

          availableSlots: cRows[i][2] || "",

          note: cRows[i][3] || ""

        });

      }

    }



    return responseJSON({ success: true, profile: profile, portfolio: portfolio, schedule: schedule, calendar: calendar });

  } catch (error) {

    return responseJSON({ success: false, message: error.toString() });

  }

}



function doPost(e) {

  try {

    let contents = "";

    if (e && e.postData && e.postData.contents) {

      contents = e.postData.contents;

    } else if (e && e.parameter && e.parameter.data) {

      contents = e.parameter.data;

    }



    if (!contents) {

      return responseJSON({ success: false, message: "Empty request payload" });

    }



    const body = JSON.parse(contents);

    const action = body.action || (e && e.parameter && e.parameter.action);



    if (action === "login") {

      return responseJSON(body.password === CONFIG.ADMIN_PASSWORD

        ? { success: true, message: "เข้าสู่ระบบสำเร็จ" }

        : { success: false, message: "รหัสผ่านไม่ถูกต้อง" });

    }



    if (body.password !== CONFIG.ADMIN_PASSWORD) {

      return responseJSON({ success: false, message: "Unauthorized: รหัสผ่านไม่ถูกต้อง" });

    }



    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

    const imageFolder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    const videoFolder = DriveApp.getFolderById(CONFIG.DRIVE_VIDEO_FOLDER_ID);



    // ----------------------------------------------------

    // Action: Save Video (บันทึกลิงก์/ID วิดีโอลง Spreadsheet)

    // ----------------------------------------------------

    if (action === "saveVideoLink") {

      const sheet = ss.getSheetByName(CONFIG.SHEET_PORTFOLIO);

      const itemId = Utilities.getUuid();

      let fileId = driveIdFrom(body.videoUrl);

      let previewUrl = body.videoUrl;



      if (fileId && fileId !== body.videoUrl) {

        previewUrl = "https\://drive.google.com/file/d/" + fileId + "/preview";

      }



      sheet.appendRow([

        itemId,

        body.title || "-",

        body.category || "General",

        "video",

        previewUrl,

        fileId,

        body.extLink || "",

        Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm:ss"),

        body.details || "",

        9999

      ]);



      return responseJSON({ success: true, message: "บันทึกผลงานวิดีโอสำเร็จ" });

    }



    if (action === "updateProfile") {

      const pSheet = ss.getSheetByName(CONFIG.SHEET_PROFILE);

      let avatarUrl = body.avatarUrl || "";

      let featuredPhotos = body.featuredPhotos || "";



      if (body.avatarData) {

        try {

          const ct = body.avatarData.substring(body.avatarData.indexOf(":") + 1, body.avatarData.indexOf(";"));

          const bytes = Utilities.base64Decode(body.avatarData.split(",")[1]);

          const blob = Utilities.newBlob(bytes, ct, "avatar\_" + Date.now());

          const file = imageFolder.createFile(blob);

          try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}

          avatarUrl = "https\://lh3.googleusercontent.com/d/" + file.getId();

        } catch (e) {}

      }



      if (body.newFeaturedFiles && body.newFeaturedFiles.length > 0) {

        let uploadedUrls = [];

        for (let f of body.newFeaturedFiles) {

          try {

            const ct = f.data.substring(f.data.indexOf(":") + 1, f.data.indexOf(";"));

            const bytes = Utilities.base64Decode(f.data.split(",")[1]);

            const blob = Utilities.newBlob(bytes, ct, "feat\_" + Date.now() + "\_" + f.name);

            const file = imageFolder.createFile(blob);

            try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}

            uploadedUrls.push("https\://lh3.googleusercontent.com/d/" + file.getId());

          } catch (e) {}

        }

        featuredPhotos = featuredPhotos ? (featuredPhotos + "," + uploadedUrls.join(",")) : uploadedUrls.join(",");

      }



      pSheet.clearContents();

      const newRows = [

        ["Key", "Value"],

        ["brandName", body.brandName || "PHOTOGRAPHER"],

        ["heroTitle", body.heroTitle || ""],

        ["heroSubtitle", body.heroSubtitle || ""],

        ["avatarUrl", avatarUrl || ""],

        ["featuredPhotos", featuredPhotos],

        ["photographerName", body.photographerName || ""],

        ["bio", body.bio || ""],

        ["location", body.location || ""],

        ["phone", body.phone || ""],

        ["line", body.line || ""],

        ["instagram", body.instagram || ""],

        ["facebook", body.facebook || ""],

        ["tiktok", body.tiktok || ""],

        ["email", body.email || ""],

        ["socialLinks", body.socialLinks || "[]"]

      ];

      pSheet.getRange(1, 1, newRows.length, 2).setValues(newRows);

      return responseJSON({ success: true, message: "บันทึกข้อมูลโปรไฟล์สำเร็จ" });

    }



    if (action === "setCalendar") {

      const cSheet = ss.getSheetByName(CONFIG.SHEET_CALENDAR);

      const rows = cSheet.getDataRange().getValues();

      let targetDateStr = body.date;

      let foundIndex = -1;



      for (let i = 1; i < rows.length; i++) {

        let rawDate = rows[i][0];

        let rowDateStr = rawDate instanceof Date

          ? Utilities.formatDate(rawDate, CONFIG.TIMEZONE, "yyyy-MM-dd")

          : String(rawDate).replace(/^'/, '');

        if (rowDateStr === targetDateStr) { foundIndex = i + 1; break; }

      }



      if (body.isDelete) {

        if (foundIndex > 0) cSheet.deleteRow(foundIndex);

        return responseJSON({ success: true, message: "รีเซ็ตสถานะวันว่างสำเร็จ" });

      }



      if (foundIndex > 0) {

        cSheet.getRange(foundIndex, 2, 1, 3).setValues([[body.statusType, body.availableSlots || "", body.note || ""]]);

      } else {

        cSheet.appendRow(["'" + targetDateStr, body.statusType, body.availableSlots || "", body.note || ""]);

      }

      return responseJSON({ success: true, message: "อัปเดตสถานะปฏิทินสำเร็จ" });

    }



    if (action === "addSchedule") {

      const sheet = ss.getSheetByName(CONFIG.SHEET_SCHEDULE);

      const itemId = body.id || Utilities.getUuid();

      const rows = sheet.getDataRange().getValues();

      let foundRow = -1;



      if (body.id) {

        for (let i = 1; i < rows.length; i++) {

          if (rows[i][0] === body.id) { foundRow = i + 1; break; }

        }

      }



      if (foundRow > 0) {

        sheet.getRange(foundRow, 2, 1, 7).setValues([[

          "'" + body.date, body.time || "ทั้งวัน", body.title || "-",

          body.location || "-", body.status || "booked", body.extLink || "", body.details || ""

        ]]);

        return responseJSON({ success: true, message: "แก้ไขคิวงานสำเร็จ" });

      } else {

        sheet.appendRow([

          itemId, "'" + body.date, body.time || "ทั้งวัน", body.title || "-",

          body.location || "-", body.status || "booked", body.extLink || "", body.details || "",

          Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm:ss")

        ]);

        return responseJSON({ success: true, message: "บันทึกคิวงานใหม่สำเร็จ" });

      }

    }



    if (action === "uploadVideoChunk") {
      const sheet = ss.getSheetByName(CONFIG.SHEET_PORTFOLIO);
      const uploadId = String(body.uploadId || "").trim();
      const chunkIndex = Number(body.chunkIndex);
      const totalChunks = Number(body.totalChunks);
      const fileName = String(body.fileName || "video.mp4");
      const mimeType = String(body.mimeType || "video/mp4");
      const chunkData = String(body.chunkData || "");
      const isLast = body.isLast === true || body.isLast === "true";

      if (!uploadId || !Number.isFinite(chunkIndex) || chunkIndex < 0 ||
          !Number.isFinite(totalChunks) || totalChunks < 1 || !chunkData) {
        return responseJSON({ success: false, message: "ข้อมูล Chunk ไม่ครบถ้วน" });
      }

      try {
        const prefix = "__VIDEO_CHUNK__" + uploadId + "__";
        const chunkName = prefix + chunkIndex;

        // ป้องกัน Chunk ซ้ำจากการ retry
        const oldChunks = videoFolder.getFilesByName(chunkName);
        while (oldChunks.hasNext()) {
          try { oldChunks.next().setTrashed(true); } catch (e) {}
        }

        const bytes = Utilities.base64Decode(chunkData);
        videoFolder.createFile(
          Utilities.newBlob(bytes, "application/octet-stream", chunkName)
        );

        if (!isLast) {
          return responseJSON({
            success: true,
            completed: false,
            chunkIndex: chunkIndex,
            message: "รับข้อมูลวิดีโอ Chunk " + (chunkIndex + 1) + "/" + totalChunks
          });
        }

        const chunkFiles = [];
        const files = videoFolder.getFiles();
        while (files.hasNext()) {
          const f = files.next();
          if (f.getName().indexOf(prefix) === 0) chunkFiles.push(f);
        }

        chunkFiles.sort(function(a, b) {
          return Number(a.getName().substring(prefix.length)) -
                 Number(b.getName().substring(prefix.length));
        });

        if (chunkFiles.length !== totalChunks) {
          return responseJSON({
            success: false,
            message: "วิดีโอส่งมาไม่ครบ Chunk (" + chunkFiles.length + "/" + totalChunks + ")"
          });
        }

        const allBytes = [];
        for (let i = 0; i < chunkFiles.length; i++) {
          const currentIndex = Number(chunkFiles[i].getName().substring(prefix.length));
          if (currentIndex !== i) {
            return responseJSON({
              success: false,
              message: "ลำดับ Chunk ของวิดีโอไม่ถูกต้อง"
            });
          }

          const part = chunkFiles[i].getBlob().getBytes();
          for (let j = 0; j < part.length; j++) allBytes.push(part[j]);
        }

        let safeName = fileName.replace(/[\\/:*?"<>|]/g, "_").trim();
        if (!safeName) safeName = "video.mp4";

        const finalFile = videoFolder.createFile(
          Utilities.newBlob(allBytes, mimeType, safeName)
        );

        try {
          finalFile.setSharing(
            DriveApp.Access.ANYONE_WITH_LINK,
            DriveApp.Permission.VIEW
          );
        } catch (shareErr) {
          Logger.log("Video sharing error: " + shareErr.toString());
        }

        const fileId = finalFile.getId();
        const previewUrl =
          "https://drive.google.com/file/d/" + fileId + "/preview";

        sheet.appendRow([
          Utilities.getUuid(),
          body.title || "-",
          body.category || "General",
          "video",
          previewUrl,
          fileId,
          body.extLink || "",
          Utilities.formatDate(
            new Date(),
            CONFIG.TIMEZONE,
            "yyyy-MM-dd HH:mm:ss"
          ),
          body.details || "",
          9999
        ]);

        chunkFiles.forEach(function(f) {
          try { f.setTrashed(true); } catch (e) {}
        });

        return responseJSON({
          success: true,
          completed: true,
          fileId: fileId,
          fileUrl: previewUrl,
          message: "อัปโหลดวิดีโอสำเร็จ"
        });

      } catch (videoErr) {
        Logger.log("Video Chunk Upload Error: " + videoErr.toString());
        return responseJSON({
          success: false,
          message: "อัปโหลดวิดีโอล้มเหลว: " + videoErr.toString()
        });
      }
    }

    if (action === "upload") {

      const sheet = ss.getSheetByName(CONFIG.SHEET_PORTFOLIO);

      const files = body.files;



      for (let f of files) {

        let directUrl = f.directUrl || "";

        let fileId = f.fileId || "";



        if (f.fileData) {

          try {

            const dataParts = f.fileData.split(",");

            const header = dataParts[0];

            const base64Data = dataParts[1];

            const mimeMatch = header.match(/:(.\*?);/);

            const ct = mimeMatch ? mimeMatch[1] : (f.mediaType === "video" ? "video/mp4" : "image/jpeg");



            const bytes = Utilities.base64Decode(base64Data);

            const originalName = (f.fileName || "file").replace(/[\/:*?"<>|]/g, "_");
            const ext = (originalName.match(/\.[^.]+$/) || [""])[0];
            const baseName = originalName.replace(/\.[^.]+$/, "") || "file";
            const fileNameUnique = baseName + "_" + Date.now() + ext;
            const blob = Utilities.newBlob(bytes, ct, fileNameUnique);
            const targetFolder = f.mediaType === "video" ? videoFolder : imageFolder;
            const file = targetFolder.createFile(blob);



            try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}

            fileId = file.getId();



            directUrl = (f.mediaType === "video")

              ? "https\://drive.google.com/file/d/" + fileId + "/preview"

              : "https\://lh3.googleusercontent.com/d/" + fileId;

          } catch (fileErr) {

            Logger.log("File Upload Error: " + fileErr.toString());
            return responseJSON({
              success: false,
              message: "อัปโหลดไฟล์ " + (f.fileName || "") +
                       " ไม่สำเร็จ: " + fileErr.toString()
            });
          }

        } else if (directUrl || fileId) {

          fileId = fileId || driveIdFrom(directUrl);

          directUrl = toDisplayUrl(directUrl, fileId, f.mediaType);

        }



        const itemId = Utilities.getUuid();

        sheet.appendRow([

          itemId,

          f.title || "-",

          f.category || "General",

          f.mediaType || "image",

          directUrl,

          fileId,

          f.extLink || "",

          Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm:ss"),

          f.details || "",

          9999

        ]);

      }

      return responseJSON({ success: true, message: "อัปโหลดผลงานสำเร็จ" });

    }



    if (action === "editPortfolioGroup") {

      const sheet = ss.getSheetByName(CONFIG.SHEET_PORTFOLIO);

      const rows = sheet.getDataRange().getValues();

      let updatedCount = 0;



      for (let i = 1; i < rows.length; i++) {

        if (rows[i][1] === body.oldTitle && rows[i][3] === body.mediaType) {

          sheet.getRange(i + 1, 2).setValue(body.newTitle);

          sheet.getRange(i + 1, 3).setValue(body.category);

          sheet.getRange(i + 1, 7).setValue(body.extLink || "");

          sheet.getRange(i + 1, 9).setValue(body.details || "");

          updatedCount++;

        }

      }

      return responseJSON({ success: true, message: `อัปเดตข้อมูลอัลบั้มสำเร็จ (${updatedCount} รายการ)` });

    }



    if (action === "reorderPortfolio") {

      const sheet = ss.getSheetByName(CONFIG.SHEET_PORTFOLIO);

      const rows = sheet.getDataRange().getValues();

      const orderMap = body.orderMap || {};



      for (let i = 1; i < rows.length; i++) {

        const title = rows[i][1];

        const mediaType = rows[i][3];

        const key = `${mediaType}\_${title}`;



        if (orderMap[key] !== undefined) {

          sheet.getRange(i + 1, 10).setValue(orderMap[key]);

        }

      }

      return responseJSON({ success: true, message: "บันทึกลำดับอัลบั้มเรียบร้อยแล้ว" });

    }



    if (action === "delete") {

      if (body.type === "portfolioGroup") {

        const sheet = ss.getSheetByName(CONFIG.SHEET_PORTFOLIO);

        const rows = sheet.getDataRange().getValues();

        for (let i = rows.length - 1; i >= 1; i--) {

          if (rows[i][1] === body.title && rows[i][3] === body.mediaType) {

            if (rows[i][5]) {

              try {

                var fDel = DriveApp.getFileById(rows[i][5]);

                if (fDel) fDel.setTrashed(true);

              } catch (e) {}

            }

            sheet.deleteRow(i + 1);

          }

        }

        return responseJSON({ success: true, message: "ลบอัลบั้มเรียบร้อยแล้ว" });

      }



      const targetSheet = body.type === "schedule" ? CONFIG.SHEET_SCHEDULE : CONFIG.SHEET_PORTFOLIO;

      const sheet = ss.getSheetByName(targetSheet);

      const rows = sheet.getDataRange().getValues();

      for (let i = 1; i < rows.length; i++) {

        if (rows[i][0] === body.id) {

          sheet.deleteRow(i + 1);

          if (rows[i][5] && body.type !== "schedule") {

            try {

              var fDel2 = DriveApp.getFileById(rows[i][5]);

              if (fDel2) fDel2.setTrashed(true);

            } catch (e) {}

          }

          return responseJSON({ success: true, message: "ลบข้อมูลสำเร็จ" });

        }

      }

      return responseJSON({ success: false, message: "ไม่พบข้อมูลที่ต้องการลบ" });

    }



    return responseJSON({ success: false, message: "คำขอไม่ถูกต้อง (Invalid action)" });

  } catch (error) {

    return responseJSON({ success: false, message: error.toString() });

  }

}