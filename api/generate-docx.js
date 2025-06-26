import { Document, Packer, Paragraph, TextRun } from "docx";

export default async function handler(req, res) {
  // CORS headers
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).send("Method Not Allowed");

  const token = req.headers.authorization;
  if (token !== "Bearer bolt-secret") return res.status(401).send("Unauthorized");

  const data = req.body;

  const createField = (label, value = "") =>
    new Paragraph({ text: `${label}: ${value}`, spacing: { after: 200 } });

  const children = [
    new Paragraph({
      children: [new TextRun({ text: "📄 Project Brief", bold: true, size: 28 })],
      spacing: { after: 300 },
    }),

    createField("Project Name", data.project_name),
    createField("Author", data.author),
    createField("Brief Added On", data.brief_added_on),
    createField("Due Date", data.due_date),
    createField("Release Date", data.release_date),
    createField("Slack Thread", data.slack_thread),

    new Paragraph({
      children: [new TextRun({ text: "\n🧠 About the Project", bold: true, size: 24 })],
    }),
    createField("About Project", data.about_project),
    createField("Target Audience", data.target_audience),
    createField("Target Outcome", data.target_outcome),

    new Paragraph({
      children: [new TextRun({ text: "\n📦 Assets & Channels", bold: true, size: 24 })],
    }),
    createField("Assets Required", data.assets_required),
    createField("Channels of Distribution", data.channels_of_distribution),
    createField("Date of Launch", data.date_of_launch),

    new Paragraph({
      children: [new TextRun({ text: "\n🔗 Links", bold: true, size: 24 })],
    }),
    createField("Important Links", data.important_links),
    createField("Creative References", data.creative_references),
    createField("Additional Info", data.additional_info),

    new Paragraph({
      children: [new TextRun({ text: "\n📅 Team Roles & Milestones", bold: true, size: 24 })],
    }),
  ];

  // Handle team_roles array
  if (Array.isArray(data.team_roles)) {
    data.team_roles.forEach((role, index) => {
      children.push(
        new Paragraph({
          children: [new TextRun({ text: `🔹 Milestone ${index + 1}`, bold: true })],
          spacing: { after: 100 },
        }),
        createField("Description", role.milestone_description),
        createField("Milestone Date", role.milestone_date),
        createField("Status", role.milestone_status),
        createField("Add Date", role.add_date),
        createField("Add Date Status", role.add_date_status)
      );
    });
  }

  const doc = new Document({
    sections: [{ children }],
  });

  const buffer = await Packer.toBuffer(doc);

  res.setHeader("Content-Disposition", "attachment; filename=brief.docx");
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
  res.send(buffer);
}
