const FOLDER_ID = '121NE9W-5Jk7B1Z566PD3AD7sHwGRxASY';
const ALERT_EMAIL = 'moseriewere@gmail.com';

function checkForNewFiles() {
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();
  const props = PropertiesService.getScriptProperties();
  const seenRaw = props.getProperty('seenFiles');
  const seen = seenRaw ? JSON.parse(seenRaw) : {};
  const newFiles = [];

  while (files.hasNext()) {
    const file = files.next();
    const id = file.getId();
    if (!seen[id]) {
      seen[id] = true;
      newFiles.push({
        name: file.getName(),
        creator: file.getOwner().getEmail(),
        date: file.getDateCreated().toLocaleString(),
        url: file.getUrl()
      });
    }
  }

  if (newFiles.length > 0) {
    const lines = newFiles.map(f =>
      `• ${f.name}\n  Added by: ${f.creator}\n  Date: ${f.date}\n  Link: ${f.url}`
    ).join('\n\n');
    MailApp.sendEmail({
      to: ALERT_EMAIL,
      subject: `${newFiles.length} new file(s) added to your folder`,
      body: `New files detected:\n\n${lines}`
    });
  }

  props.setProperty('seenFiles', JSON.stringify(seen));
}
