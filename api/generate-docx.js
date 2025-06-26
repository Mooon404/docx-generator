import { Document, Packer, Paragraph, TextRun } from "docx";

export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).send("Method Not Allowed");

  // Optional: Secure the endpoint
  const token = req.headers.authorization;
  if (token !== "Bearer YOUR_SECRET_TOKEN") {
    return res.status(401).send("Unauthorized");
  }

  const data = req.body;

  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({ children: [new TextRun({ text: "Project Brief", bold: true, size: 28 })] }),
          new Paragraph({ text: `Project Name: ${data.project_name || ''}` }),
          new Paragraph({ text: `Author: ${data.author || ''}` }),
          new Paragraph({ text: `Due Date: ${data.due_date || ''}` }),
          new Paragraph({ text: `About Project: ${data.about_project || ''}` }),
          new Paragraph({ text: `Target Audience: ${data.target_audience || ''}` }),
          new Paragraph({ text: `Assets Required: ${data.assets_required || ''}` })
        ]
      }
    ]
  });

  const buffer = await Packer.toBuffer(doc);
  res.setHeader("Content-Disposition", "attachment; filename=brief.docx");
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
  res.send(buffer);
}
