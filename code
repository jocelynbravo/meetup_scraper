// ============================================================
// FULL TECHLAHOMA SCRIPT — Replace entire contents with this
// ============================================================

function syncTechlahomaEvents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const today = new Date();
  const todayStr = Utilities.formatDate(today, tz, "yyyy-MM-dd");

  let upcomingSheet = ss.getSheetByName("Upcoming Events");
  let archiveSheet  = ss.getSheetByName("All Events");

  if (!upcomingSheet) upcomingSheet = ss.insertSheet("Upcoming Events");
  if (!archiveSheet)  archiveSheet  = ss.insertSheet("All Events");

  const headers = [
    "Event ID","Event Name","Group","Date","Venue","City",
    "Link","Status","First Seen","Last Seen","Image URL","Address","Description"
  ];

  if (upcomingSheet.getLastRow() === 0) upcomingSheet.appendRow(headers);
  if (archiveSheet.getLastRow() === 0)  archiveSheet.appendRow(headers);

  const events = fetchTechlahomaEvents();
  events.sort((a, b) => new Date(a.dateTime) - new Date(b.dateTime));

  const upcomingData = upcomingSheet.getDataRange().getValues();
  const upcomingMap = {};
  for (let i = 1; i < upcomingData.length; i++) upcomingMap[upcomingData[i][0]] = i + 1;

  const archiveData = archiveSheet.getDataRange().getValues();
  const archiveMap = {};
  for (let i = 1; i < archiveData.length; i++) archiveMap[archiveData[i][0]] = i + 1;

  const seen = {};

  events.forEach(e => {
    const id       = e.eventUrl;
    seen[id]       = true;
    const dateObj  = new Date(e.dateTime);
    const date     = Utilities.formatDate(dateObj, tz, "yyyy-MM-dd HH:mm");
    const venue    = e.venue?.name && e.venue.name !== "Online event" ? e.venue.name : "";
    const city     = e.venue?.city || e.group?.city || "";
    const group    = e.group?.name || "";
    const imageUrl = e.featuredEventPhoto?.highResUrl || "";
    const address  = (e.venue?.address && e.venue.address.trim() !== "")
                     ? e.venue.address + (city ? ", " + city : "")
                     : "";
    const description = stripMarkdown(e.description || "");

    const row = [
      id, e.title.trim(), group, date, venue, city,
      e.eventUrl, "Active", todayStr, todayStr, imageUrl, address, description
    ];

    if (archiveMap[id]) {
      const r = archiveMap[id];
      archiveSheet.getRange(r, 2, 1, 6).setValues([row.slice(1, 7)]);
      archiveSheet.getRange(r, 8).setValue("Active");
      archiveSheet.getRange(r, 10).setValue(todayStr);
      archiveSheet.getRange(r, 11).setValue(imageUrl);
      archiveSheet.getRange(r, 12).setValue(address);
      archiveSheet.getRange(r, 13).setValue(description);
    } else {
      archiveSheet.appendRow(row);
      archiveMap[id] = archiveSheet.getLastRow();
    }

    if (dateObj >= today) {
      if (upcomingMap[id]) {
        const r = upcomingMap[id];
        upcomingSheet.getRange(r, 2, 1, 6).setValues([row.slice(1, 7)]);
        upcomingSheet.getRange(r, 8).setValue("Active");
        upcomingSheet.getRange(r, 10).setValue(todayStr);
        upcomingSheet.getRange(r, 11).setValue(imageUrl);
        upcomingSheet.getRange(r, 12).setValue(address);
        upcomingSheet.getRange(r, 13).setValue(description);
      } else {
        upcomingSheet.appendRow(row);
        upcomingMap[id] = upcomingSheet.getLastRow();
      }
    }
  });

  // Mark removed events in archive
  const updatedArchive = archiveSheet.getDataRange().getValues();
  for (let i = 1; i < updatedArchive.length; i++) {
    const id = updatedArchive[i][0];
    if (id && !seen[id]) archiveSheet.getRange(i + 1, 8).setValue("Removed from Meetup");
  }

  // Remove past events from upcoming
  const updatedUpcoming = upcomingSheet.getDataRange().getValues();
  for (let i = updatedUpcoming.length - 1; i >= 1; i--) {
    if (new Date(updatedUpcoming[i][3]) < today) upcomingSheet.deleteRow(i + 1);
  }

  // Sort both sheets
  if (upcomingSheet.getLastRow() > 1)
    upcomingSheet.getRange(2, 1, upcomingSheet.getLastRow() - 1, 13).sort({ column: 4, ascending: true });
  if (archiveSheet.getLastRow() > 1)
    archiveSheet.getRange(2, 1, archiveSheet.getLastRow() - 1, 13).sort({ column: 4, ascending: true });

  Logger.log("Sync complete.");
}

// Strip markdown syntax, preserve line breaks
function stripMarkdown(text) {
  return text
    .replace(/^#{1,6}\s+/gm, "")
    .replace(/\*\*(.*?)\*\*/g, "$1")
    .replace(/\*(.*?)\*/g, "$1")
    .replace(/__(.*?)__/g, "$1")
    .replace(/_(.*?)_/g, "$1")
    .replace(/\[([^\]]+)\]\([^)]+\)/g, "$1")
    .replace(/`([^`]+)`/g, "$1")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function fetchTechlahomaEvents() {
  const payload = {
    operationName: "getProNetworkEventsByUrlname",
    variables: { urlname: "techlahoma", first: 50 },
    query: `
      query getProNetworkEventsByUrlname($urlname: ID!, $first: Int) {
        proNetwork(urlname: $urlname) {
          suggestedEvents(input: { first: $first, filter: { latitude: 35.48, longitude: -97.53 } }) {
            edges {
              node {
                id
                title
                dateTime
                eventUrl
                description
                featuredEventPhoto { highResUrl }
                group { name city urlname }
                venue { name city lat lon address }
              }
            }
          }
        }
      }
    `
  };

  try {
    const response = UrlFetchApp.fetch("https://www.meetup.com/gql2", {
      method: "post",
      contentType: "application/json",
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "application/json",
        "Origin": "https://www.meetup.com",
        "Referer": "https://www.meetup.com/pro/techlahoma/"
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const json = JSON.parse(response.getContentText());
    const edges = json?.data?.proNetwork?.suggestedEvents?.edges || [];
    Logger.log("Events returned: " + edges.length);
    return edges.map(e => e.node);
  } catch(err) {
    Logger.log("Error: " + err.message);
    return [];
  }
}

// Strip HTML tags from Meetup description, preserve line breaks
function stripHtml(html) {
  return html
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function generateNewsletterDoc() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Upcoming Events");
  const data = sheet.getDataRange().getValues();
  const tz = "America/Chicago";

  const groups = { "Oklahoma City": [], "Tulsa": [], "Virtual": [] };

  for (let i = 1; i < data.length; i++) {
    const row         = data[i];
    const link        = row[0];
    const title       = row[1];
    const group       = row[2];
    const rawDate     = row[3];
    const venue       = row[4] || "";
    const city        = (row[5] || "").toLowerCase();
    const imageUrl    = row[10] || "";
    const address     = row[11] || "";
    const description = row[12] || "";

    let dateObj = rawDate instanceof Date ? rawDate : new Date(rawDate);
    const dateStr = Utilities.formatDate(dateObj, tz, "EEEE, MMMM d @ h:mm a") + " CST";

    const event = { link, title, group, dateStr, venue, city: row[5] || "", imageUrl, address, description };

    const isVirtual = !venue || venue.toLowerCase() === "online event" || venue.trim() === "";
    const isTulsa   = city.includes("tulsa");
    const isOKC     = city.includes("oklahoma city") || city.includes("okc");

    if (isVirtual)    groups["Virtual"].push(event);
    else if (isTulsa) groups["Tulsa"].push(event);
    else if (isOKC)   groups["Oklahoma City"].push(event);
    else groups[group.toLowerCase().includes("tulsa") ? "Tulsa" : "Oklahoma City"].push(event);
  }

  // Save to same folder as spreadsheet
  const docName = "Techlahoma Newsletter Events – " + Utilities.formatDate(new Date(), tz, "MMMM d, yyyy");
  const doc = DocumentApp.create(docName);
  const docFile = DriveApp.getFileById(doc.getId());
  const ssFile = DriveApp.getFileById(ss.getId());
  const parents = ssFile.getParents();
  if (parents.hasNext()) {
    const targetFolder = parents.next();
    targetFolder.addFile(docFile);
    DriveApp.getRootFolder().removeFile(docFile);
  }

  const body = doc.getBody();
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(54).setMarginRight(54);

  // Title
  const titlePara = body.appendParagraph("🖥️ Techlahoma Upcoming Events");
  titlePara.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  titlePara.editAsText().setFontSize(22).setBold(true).setForegroundColor("#0f1624");

  const subtitlePara = body.appendParagraph(
    "Powered by Techlahoma Foundation · Oklahoma's Largest Technology Network"
  );
  subtitlePara.editAsText().setFontSize(10).setItalic(true).setForegroundColor("#7a8fa6");
  body.appendParagraph("");

  const sectionOrder = ["Oklahoma City", "Tulsa", "Virtual"];
  const sectionIcons = { "Oklahoma City": "🏙️", "Tulsa": "🌆", "Virtual": "💻" };

  sectionOrder.forEach(sectionName => {
    const events = groups[sectionName];
    if (events.length === 0) return;

    // Section header
    const header = body.appendParagraph(sectionIcons[sectionName] + " " + sectionName);
    header.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    header.editAsText().setFontSize(14).setBold(true).setForegroundColor("#00b4d8");

    events.forEach(e => {

      if (e.imageUrl) {
        try {
          const imgBlob = UrlFetchApp.fetch(e.imageUrl).getBlob();
          const img = body.appendImage(imgBlob);
          const maxWidth = 300;  // Adjust this to taste (points, ~4 inches)
          const w = img.getWidth();
          const h = img.getHeight();
          const scale = maxWidth / w;
          img.setWidth(maxWidth);
          img.setHeight(Math.round(h * scale));
          body.appendParagraph("");
        } catch(err) {
          Logger.log("Image fetch failed for: " + e.title + " — " + err.message);
        }
      }

      // Group name
      body.appendParagraph(e.group)
        .editAsText().setFontSize(12).setBold(true).setForegroundColor("#111111");

      // Event title — linked
      body.appendParagraph(e.title)
        .editAsText().setFontSize(12).setBold(false).setForegroundColor("#00b4d8").setLinkUrl(e.link);

      // Date
      body.appendParagraph(e.dateStr)
        .editAsText().setFontSize(11).setBold(true).setForegroundColor("#222222");

      // Venue name (if not virtual)
      if (e.venue && e.venue.toLowerCase() !== "online event") {
        body.appendParagraph(e.venue)
          .editAsText().setFontSize(11).setBold(false).setForegroundColor("#333333");
      }

      // Address — linked to Google Maps
      if (e.address) {
        const mapsUrl = "https://www.google.com/maps/search/?api=1&query=" +
          encodeURIComponent(e.address);
        body.appendParagraph(e.address)
          .editAsText().setFontSize(11).setForegroundColor("#00b4d8").setLinkUrl(mapsUrl);
      }

      // Divider
      body.appendParagraph("──────────────────────────────────")
        .editAsText().setFontSize(8).setForegroundColor("#cccccc");

      // Description
      if (e.description) {
        body.appendParagraph(e.description)
          .editAsText().setFontSize(11).setForegroundColor("#444444").setItalic(false);
        body.appendParagraph("");
      }

    });

    body.appendParagraph("");
  });

  // Footer
  body.appendParagraph(
    "Last updated: " + Utilities.formatDate(new Date(), tz, "MMMM d, yyyy h:mm a") + " CST"
  ).editAsText().setFontSize(9).setItalic(true).setForegroundColor("#999999");

  doc.saveAndClose();
  Logger.log("Newsletter doc created: " + doc.getUrl());

  // Send email to recipients
  const newsletterRecipients = [
    Session.getActiveUser().getEmail(),
    "jocelyn_bravo@hotmail.com",
    
  ].join(",");

  Logger.log("Sending to: " + newsletterRecipients);

  GmailApp.sendEmail(
    newsletterRecipients,
    "Techlahoma Newsletter Doc Ready",
    "Your newsletter draft is ready:\n\n" + doc.getUrl()
  );
}

function checkForNewGroups() {
  const knownSlugs = new Set([
    "oklahoma-city-java-user-group","oklahoma-city-techlahoma","okcwebdevs",
    "okccoffeeandcode","okc-osh","oklahom_ai","okc-lugnuts",
    "oklahoma-game-developers","okc-sharp","salesforce-meetup-group",
    "pythonistas","tulsadevelopers-net","tulsa-ux-user-group",
    "tulsa-web-devs","techlahoma-foundation-tulsa","tulsaagilepractitioners",
    "oklahomai-developers","tulsa-aws","tulsa-game-developers"
  ]);

  const html = UrlFetchApp.fetch("https://www.meetup.com/pro/techlahoma/", {
    headers: { "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36" }
  }).getContentText();

  const matches = html.match(/meetup\.com\/([a-zA-Z0-9_-]+)\/events/g) || [];
  const found = [...new Set(matches.map(m => m.match(/meetup\.com\/([a-zA-Z0-9_-]+)\//)[1]))];
  const newGroups = found.filter(s => !knownSlugs.has(s.toLowerCase()));

  if (newGroups.length > 0) {
    Logger.log("NEW GROUPS FOUND: " + newGroups.join(", "));
    GmailApp.sendEmail(
      Session.getActiveUser().getEmail(),
      "New Techlahoma Groups Found",
      "New groups detected on meetup.com/pro/techlahoma:\n\n" + newGroups.join("\n")
    );
  } else {
    Logger.log("No new groups found.");
  }
}

function createWeeklyTrigger() {
  ScriptApp.newTrigger("checkForNewGroups")
    .timeBased().everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).create();

  ScriptApp.newTrigger("syncTechlahomaEvents")
    .timeBased().everyDays(1).atHour(6).create();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📅 Techlahoma")
    .addItem("Sync Events", "syncTechlahomaEvents")
    .addItem("Generate Newsletter Doc", "generateNewsletterDoc")
    .addSeparator()
    .addItem("Check for New Groups", "checkForNewGroups")
    .addToUi();
}

function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Upcoming Events");
  const data = sheet.getDataRange().getValues();
  const tz = "America/Chicago";

  const groups = { "Oklahoma City": [], "Tulsa": [], "Virtual": [] };

  for (let i = 1; i < data.length; i++) {
    const row     = data[i];
    const link    = row[0];
    const title   = row[1];
    const group   = row[2];
    const rawDate = row[3];
    const venue   = row[4] || "";
    const city    = (row[5] || "").toLowerCase();
    const address = row[11] || "";

    let dateObj;
    if (rawDate instanceof Date) {
      dateObj = rawDate;
    } else {
      dateObj = new Date(rawDate);
    }

    const dateStr = Utilities.formatDate(dateObj, tz, "EEEE, MMMM d, yyyy");
    const timeStr = Utilities.formatDate(dateObj, tz, "h:mm a") + " CST";

    const event = { link, title, group, dateStr, timeStr, venue, city: row[5] || "", address };

    const isVirtual = !venue || venue.toLowerCase() === "online event" || venue.trim() === "";
    const isTulsa   = city.includes("tulsa");
    const isOKC     = city.includes("oklahoma city") || city.includes("okc");

    if (isVirtual)    groups["Virtual"].push(event);
    else if (isTulsa) groups["Tulsa"].push(event);
    else if (isOKC)   groups["Oklahoma City"].push(event);
    else groups[group.toLowerCase().includes("tulsa") ? "Tulsa" : "Oklahoma City"].push(event);
  }

  function buildSection(icon, label, events) {
    if (events.length === 0) return "";
    let text = `${icon} *${label}*\n\n`;
    events.forEach(e => {
      const location = e.venue && e.venue.toLowerCase() !== "online event"
        ? (e.address ? e.address : e.venue + (e.city ? ", " + e.city : ""))
        : "Online";

      const cleanTitle = e.title.replace(/[|&<>]/g, m => ({
        '|': '-', '&': 'and', '<': '', '>': ''
      }[m]));

      text += `*${cleanTitle}*\n`;
      text += `_${e.group}_\n`;
      text += `🔗 ${e.link}\n`;
      text += `📅 ${e.dateStr}\n`;
      text += `📍 ${location}\n\n`;
    });
    return text + "---\n\n";
  }

  const sectionsHtml =
    buildSection("Oklahoma City", "🏙️", groups["Oklahoma City"]) +
    buildSection("Tulsa", "🌆", groups["Tulsa"]) +
    buildSection("Virtual", "💻", groups["Virtual"]);

  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <title>Techlahoma Upcoming Events</title>
      <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: Arial, sans-serif; background: #0f1624; color: #e0e0e0; padding: 40px 20px; }
        .container { max-width: 960px; margin: 0 auto; }
        h1 { font-size: 28px; color: #ffffff; font-weight: bold; margin-bottom: 6px; }
        h1 span { color: #00b4d8; }
        .subtitle { color: #7a8fa6; font-size: 14px; margin-bottom: 36px; }
        .subtitle a { color: #00b4d8; text-decoration: none; }
        .subtitle a:hover { text-decoration: underline; }
        .section { margin-bottom: 40px; }
        .section h2 { font-size: 16px; color: #00b4d8; margin-bottom: 12px; padding-bottom: 8px;
          border-bottom: 2px solid #1a2235; text-transform: uppercase; letter-spacing: 1px; }
        table { width: 100%; border-collapse: collapse; background: #1a2235;
          border-radius: 10px; overflow: hidden; box-shadow: 0 4px 20px rgba(0,0,0,0.4); }
        th { background: #00b4d8; color: #0f1624; padding: 12px 16px; text-align: left;
          font-size: 12px; font-weight: bold; text-transform: uppercase; letter-spacing: 0.5px; }
        td { padding: 12px 16px; border-bottom: 1px solid #243050; font-size: 14px;
          vertical-align: top; color: #c8d6e5; }
        tr:last-child td { border-bottom: none; }
        tr:hover td { background: #1f2d45; }
        a { color: #00b4d8; text-decoration: none; font-weight: 600; }
        a:hover { text-decoration: underline; }
        .date { font-weight: 600; color: #ffffff; }
        .time { font-size: 12px; color: #7a8fa6; margin-top: 3px; }
        .updated { margin-top: 18px; font-size: 12px; color: #4a5f7a; text-align: right; }
        @media (max-width: 600px) { th, td { padding: 10px; font-size: 12px; } h1 { font-size: 20px; } }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>🖥️ <span>Techlahoma</span> Upcoming Events</h1>
        <p class="subtitle">
          Powered by <a href="https://www.techlahoma.org" target="_blank">Techlahoma Foundation</a>
          &nbsp;·&nbsp; Oklahoma's Largest Technology Network
        </p>
        ${sectionsHtml}
        <p class="updated">Last updated: ${Utilities.formatDate(new Date(), tz, "MMMM d, yyyy h:mm a")} CST</p>
      </div>
    </body>
    </html>
  `;

  return HtmlService.createHtmlOutput(html)
    .setTitle("Techlahoma Upcoming Events")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function generateSlackPost() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Upcoming Events");
  const data = sheet.getDataRange().getValues();
  const tz = "America/Chicago";

  const groups = { "Oklahoma City": [], "Tulsa": [], "Virtual": [] };

  for (let i = 1; i < data.length; i++) {
    const row     = data[i];
    const link    = row[0];
    const title   = row[1];
    const group   = row[2];
    const rawDate = row[3];
    const venue   = row[4] || "";
    const city    = (row[5] || "").toLowerCase();
    const address = row[11] || "";

    let dateObj = rawDate instanceof Date ? rawDate : new Date(rawDate);
    const dateStr = Utilities.formatDate(dateObj, tz, "EEEE, MMMM d @ h:mm a") + " CST";

    const event = { link, title, group, dateStr, venue, city: row[5] || "", address };

    const isVirtual = !venue || venue.toLowerCase() === "online event" || venue.trim() === "";
    const isTulsa   = city.includes("tulsa");
    const isOKC     = city.includes("oklahoma city") || city.includes("okc");

    if (isVirtual)    groups["Virtual"].push(event);
    else if (isTulsa) groups["Tulsa"].push(event);
    else if (isOKC)   groups["Oklahoma City"].push(event);
    else groups[group.toLowerCase().includes("tulsa") ? "Tulsa" : "Oklahoma City"].push(event);
  }

  function buildSection(icon, label, events) {
    if (events.length === 0) return "";
    let text = icon + " *" + label + "*\n\n";
    events.forEach(e => {
      const location = e.venue && e.venue.toLowerCase() !== "online event"
        ? (e.address ? e.address : e.venue + (e.city ? ", " + e.city : ""))
        : "Online";

      text += "*" + e.title.trim() + "*\n";
      text += "_" + e.group + "_\n";
      text += e.link + "\n";
      text += "📅 " + e.dateStr + "\n";
      text += "📍 " + location + "\n\n";
    });
    return text + "---\n\n";
  }

  const post =
    buildSection(":cityscape:", "Oklahoma City", groups["Oklahoma City"]) +
    buildSection(":city_sunset:", "Tulsa", groups["Tulsa"]) +
    buildSection(":computer:", "Virtual", groups["Virtual"]);

  // Save to same folder as spreadsheet
  const fileName = "Techlahoma Slack Post – " + Utilities.formatDate(new Date(), tz, "MMMM d, yyyy");
  const file = DriveApp.createFile(fileName + ".txt", post, MimeType.PLAIN_TEXT);
  const ssFile = DriveApp.getFileById(ss.getId());
  const parents = ssFile.getParents();
  if (parents.hasNext()) {
    const folder = parents.next();
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
  }

  Logger.log("Slack post:\n\n" + post);

  GmailApp.sendEmail(
    Session.getActiveUser().getEmail(),
    "Techlahoma Slack Post Ready",
    "Copy and paste this into Slack:\n\n" + post
  );
}
