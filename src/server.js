import express from "express";
import multer from "multer";
import cors from "cors";
import XLSX from "xlsx";
import fs from "fs";

const app = express();
app.use(cors());

// 📁 Upload config
const upload = multer({ dest: "uploads/" });


// 🔹 Read Excel
const readExcel = (filePath) => {
  const wb = XLSX.readFile(filePath);
  const sheet = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet);
};


// 🔴 OVERWRITE LOGIC
const overwriteData = (master, mapping) => {
  const map = new Map();

  // master base
  master.forEach(item => {
    map.set(Number(item.ID), { ...item });
  });

  // mapping apply
  mapping.forEach(item => {
    let ids = String(item.ID).split(";");

    ids.forEach(id => {
      id = Number(id.trim());

      if (map.has(id)) {
        map.set(id, {
          ...map.get(id),
          ...item,
          ID: id
        });
      } else {
        map.set(id, {
          ...item,
          ID: id
        });
      }
    });
  });

  return Array.from(map.values());
};


// 🟣 MERGE LOGIC
const mergeData = (master, mapping, delimiter = "//") => {
  const map = new Map();

  master.forEach(item => {
    map.set(Number(item.ID), { ...item });
  });

  mapping.forEach(item => {
    let ids = String(item.ID).split(";");

    ids.forEach(id => {
      id = Number(id.trim());

      if (map.has(id)) {
        const existing = map.get(id);

        Object.keys(item).forEach(key => {
          if (key === "ID") return;

          if (existing[key] && item[key] && existing[key] !== item[key]) {
            existing[key] = `${existing[key]} ${delimiter} ${item[key]}`;
          } else if (!existing[key]) {
            existing[key] = item[key];
          }
        });

        map.set(id, existing);

      } else {
        map.set(id, {
          ...item,
          ID: id
        });
      }
    });
  });

  return Array.from(map.values());
};


// 🔹 API
app.post(
  "/upload",
  upload.fields([
    { name: "master" },
    { name: "mapping" }
  ]),
  (req, res) => {
    try {
      const type = req.body.type || "overwrite"; // 🔥 merge / overwrite
      const delimiter = req.body.delimiter || "//";

      const masterFile = req.files.master[0].path;
      const mappingFile = req.files.mapping[0].path;

      const master = readExcel(masterFile);
      const mapping = readExcel(mappingFile);

      let result;

      if (type === "merge") {
        result = mergeData(master, mapping, delimiter);
      } else {
        result = overwriteData(master, mapping);
      }

      // 📊 Convert to Excel
      const ws = XLSX.utils.json_to_sheet(result);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Output");

      const buffer = XLSX.write(wb, {
        type: "buffer",
        bookType: "xlsx",
      });

      res.setHeader(
        "Content-Disposition",
        "attachment; filename=output.xlsx"
      );
      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );

      res.send(buffer);

      // 🧹 cleanup files
      fs.unlinkSync(masterFile);
      fs.unlinkSync(mappingFile);

    } catch (err) {
      console.error(err);
      res.status(500).send("Server Error");
    }
  }
);


// 🚀 Start server
app.listen(5000, () => {
  console.log("✅ Server running on http://localhost:5000");
});