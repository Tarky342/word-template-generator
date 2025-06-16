const express = require("express");
const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const app = express();
app.use(express.json());

app.post("/generate", (req, res) => {
  const data = req.body;
  console.log("📥 リクエスト受信:", data);

  // テンプレートファイルのパス
  const templatePath = path.join(__dirname, "templates", "kyoshitsu.docx");

  try {
    console.log("📄 テンプレート読み込み中:", templatePath);
    const content = fs.readFileSync(templatePath, "binary");

    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });

    doc.setData({
      applicant: data.applicant ?? "",
      affiliation: data.affiliation ?? "",
      date: data.date ?? "",
      room: data.room ?? "",
      purpose: data.purpose ?? "",
      note: data.note ?? "",
    });

    try {
      doc.render();
      console.log("✅ テンプレート埋め込み成功");
    } catch (renderError) {
      console.error("🔥 テンプレート埋め込みエラー:", renderError);
      return res.status(500).send("テンプレート処理に失敗しました");
    }

    const buf = doc.getZip().generate({ type: "nodebuffer" });
    const outputPath = path.join(
      __dirname,
      "output",
      `kyoshitsu_${Date.now()}.docx`
    );
    fs.writeFileSync(outputPath, buf);

    console.log("📤 出力ファイル作成済:", outputPath);

    res.download(outputPath, (err) => {
      if (err) {
        console.error("❌ ダウンロードエラー:", err.message);
      } else {
        console.log("📦 ダウンロード開始:", outputPath);
      }
    });
  } catch (outerError) {
    console.error("🚨 重大エラー:", outerError);
    res.status(500).send("サーバー処理に失敗しました");
  }
});

app.listen(3000, () => {
  console.log("🌐 サーバー起動済: http://localhost:3000");
});
