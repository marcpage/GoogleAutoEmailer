/**
 * Automate sending draft emails to tagged contacts.
 * 
 * Create a draft email and put To: Tag1,Tag2,Tag3 as the first line
 *  To send on or after a date, add: Date: YYYY/MM/DD
 * 
 */

// In the Automation sheet, add a Form called "Opt Out" 
// These are from the url you get when you click the eye (view) on the form.
var OptOutFormId = "d/e/1FAIpQLSe5ZX_Yk8hVxZJSIBsNbORRHBgVe3jvvkadMGG-GwujTJZjJw";
var OptOutEmailId = "entry.1967723952"; // inspect the form page and search for "entry."
var OptOutEmailPrefix = "https://docs.google.com/forms/"+OptOutFormId+"/viewform?"+OptOutEmailId+"=";
var OptOutStyle = "font-size:small;color:gray";
var OptOutLinkStyle = "color:gray";
var RecipientsPattern = /\bTo:\s*([^<\r\n]+)/;
var SendDatePattern = /\bDate: (([0-9][0-9][0-9][0-9])\/([0-9][0-9])\/([0-9][0-9]))/;

function getEveryGroup() {
  var groups = People.ContactGroups.list();
  var group_mappings = {};

  // TODO: Use pageToken and nextPageToken for more than 100 groups
  for (var group_index = 0; group_index < groups.contactGroups.length; ++group_index) {
    group = groups.contactGroups[group_index];
    group_mappings[group.resourceName] = group.name;
  }
  return group_mappings;
}

function getEveryone() {
  var groups = getEveryGroup();
  var fields = {"personFields": "names,emailAddresses,memberships,biographies"};
  var everyone = {};
  var people; 

  do {
    people = People.People.Connections.list("people/me", fields);
    for (var person_index=0; person_index < people.connections.length; ++person_index) {
      person = people.connections[person_index];
      emails = person.emailAddresses 
                ? person.emailAddresses.map(function (address) {return address.value}) 
                : [];
      membership = person.memberships 
                ? person.memberships.map(function (group) {return groups[group.contactGroupMembership.contactGroupResourceName]}) 
                : []
      everyone[person.resourceName] = {
          "first": person.names[0].givenName,
          "last": person.names[0].familyName,
          "emails": emails,
          "groups": membership,
      }

      for (var biography_index = 0; biography_index < (person.biographies ? person.biographies.length : 0); ++ biography_index) {
        extras = person.biographies[biography_index].value.match(/\b(\S+):\s*([^<\r\n]+)/g);

        for (var variable_index = 0; variable_index < (extras ? extras.length : 0); ++variable_index) {
          var key_value = extras[variable_index].match(/\b(\S+):\s*([^<\r\n]+)/);

          everyone[person.resourceName][key_value[1]] = key_value[2];
          //console.log(everyone[person.resourceName]);
        }

      }

    }
    fields.pageToken = people.nextPageToken;
  } while(people.nextPageToken);

  return everyone;
}

function test_getPeopleWithTag() {
  var all_contacts = getEveryone();
  var people = getPeopleWithTag("Coaching Clients", all_contacts);

  console.log(people);
}
function getPeopleWithTag(tagName, all_contacts) {
  var contact_ids = Object.keys(all_contacts);
  var people = contact_ids.map(function(person_id) {return all_contacts[person_id]});

  return people.filter(function(person) {return person.groups.includes(tagName)});
}

function getDateString() {
  var now = new Date();
  var y = now.getFullYear();
  var m = now.getMonth() + 1;
  var d = now.getDate();
  var h = now.getHours();
  var min = now.getMinutes();
  var s = now.getSeconds();
  var ms = now.getMilliseconds();
  var tz = now.getTimezoneOffset();

  return y + "/"
          + (m < 10 ? '0' : '') + m + "/"
          + (d < 10 ? '0' : '') + d + " "
          + (h < 10 ? '0' : '') + h + ":"
          + (min < 10 ? '0' : '') + min + ":"
          + (s < 10 ? '0' : '') + s;
}

function initializeAutomatedEmailTrackerTab(sheet, name, subject) {
    var data = sheet.getRange("A1:C1");
    data.setValues([["email", "sent", subject]]);
    sheet.setName(name);
}

function getAutomatedEmailTracker(emailId, subject) {
  var settingsFileName = "Automated Emails";
  var documents = DriveApp.getFilesByName(settingsFileName);
  var foundDocument = undefined;
  var hasMore = documents.hasNext();
  
  while (documents.hasNext()) {
    var document = documents.next();
    if (document.getMimeType() == MimeType.GOOGLE_SHEETS) {
      foundDocument = SpreadsheetApp.open(document);
    }
  }
  
  var foundSheet = undefined;
  var optOutSheet = undefined;

  if (!foundDocument) {
    var foundDocument = SpreadsheetApp.create(settingsFileName);
    
    foundSheet = foundDocument.getSheets()[0];
    initializeAutomatedEmailTrackerTab(foundSheet, emailId, subject);
  } else {
    var sheets = foundDocument.getSheets();

    for (var index = 0; index < sheets.length; index++) {
      var sheetName = sheets[index].getName();
      if (sheetName == emailId) {
        foundSheet = sheets[index];
      }
      if (sheetName == "Opt Out") {
        optOutSheet = sheets[index];
      }
    }

    if (!foundSheet) {
      foundSheet = foundDocument.insertSheet();
      initializeAutomatedEmailTrackerTab(foundSheet, emailId, subject);
    }
  }
  
  return {'send_sheet': foundSheet, 'opt_out_sheet': optOutSheet};
}

function automatedEmailsAvailable(sheet, optOutSheet, emails) {
  var data = sheet.getRange("A2:B").getValues();
  var optOut = optOutSheet ? optOutSheet.getRange("B2:B").getValues() : undefined;
  var notAvailable = [];
  var optOutList = [];

  for (var row = 0; row < optOut.length; row++) {
    if (optOut[row] && optOut[row][0]) {
      optOutList.push(optOut[row][0]);
    }
  }

  for (var row = 0; row < data.length; row++) {
    if (emails.indexOf(data[row][0]) >= 0) {
      notAvailable.push(data[row][0]);
    }
  }
  return emails.filter(function(e) {return notAvailable.indexOf(e) < 0 && optOutList.indexOf(e) < 0;})
}

function test_getEmailList() {
  var list = getEmailList(['Coaching Clients'], getEveryone());
  console.log(list);
}

function getEmailList(recipients, all_contacts) {
  var emails = [];

  for (var recipientIndex = 0; recipientIndex < recipients.length; ++recipientIndex) {
    var people = getPeopleWithTag(recipients[recipientIndex].replace(/^\s+|\s+$/g, ''), all_contacts);
    
    Array.prototype.push.apply(emails, people);
  }
  return emails;
}

function matchAttachmentsFromBody(inlineImages, body) {
  var name_image = {};
  var inlinePattern = /<img[^>]+data-surl="cid:([^"]+)" src="cid:([^"]+)" alt="([^"]+)"[^>]+>/g;
  var inlinedImageInfo = [...body.matchAll(inlinePattern)];

  for (var inlinedIndex = 0; inlinedIndex < inlinedImageInfo.length; ++inlinedIndex) {
    if (inlinedImageInfo[inlinedIndex][1] != inlinedImageInfo[inlinedIndex][2]) {
      throw "Mismatched identifiers";
    }
  }

  if (inlinedImageInfo.length != inlineImages.length) {
    throw "Mismatched inlined image counts";
  }

  for (var imagesIndex = 0; imagesIndex < inlineImages.length; ++imagesIndex) {
    var image = inlineImages[imagesIndex];
    var imageInfo = inlinedImageInfo[imagesIndex];

    if (imageInfo[3] != image.getName()) {
      throw "Image name mismatch: " + imageInfo[3] + " vs " + image.getName();
    }

    name_image[imageInfo[1]] = image;
  }

  return name_image;
}

function markRecipients(trackingSheet, recipientEmails) {
  var data = trackingSheet.getRange("A2:B");
  var values = data.getValues();
  var emails = [...recipientEmails];

  for (var row = 0; row < values.length; row++) {
    if (!values[row][0]) {
      values[row][0] = emails.pop();
      values[row][1] = getDateString();
      if (!emails.length) {
        break;
      }
    }
  }
  data.setValues(values);
}

function replaceVariables(text, recipient, emails, isHtml) {
  var variablePattern = /\{\{([A-Za-z0-9]+)\}\}/g;
  var replacements = {};
  var instances = [...text.matchAll(variablePattern)];

  for (var instanceIndex = 0; instanceIndex < instances.length; ++instanceIndex) {
    var key = instances[instanceIndex][0];
    var value = recipient[instances[instanceIndex][1]];
    replacements[key] = value ? value : ""; 
  }
  for (var key in replacements) {
    text = text.replace(key, replacements[key]);
  }
  if (isHtml) {
    text = text + "<br/><br/><center><div style='" + OptOutStyle + "'>"
                + "If you no longer with to receive these emails you may " 
                + "<a href='" 
                  + OptOutEmailPrefix + emails[0] 
                  + "' style='" + OptOutLinkStyle + "'>Unsubscribe</a>."
                + "</div></center>"
  } else {
    text = text + "If you no longer with to receive these emails you may "
                + "Unsubscribe by visiting this web page: "
                + OptOutEmailPrefix + emails[0]
  }
  return text;
}

function sendMessage(draft, recipient, trackingSheet, optOutSheet) {
  var recipientEmails = automatedEmailsAvailable(trackingSheet, optOutSheet, recipient["emails"]);

  if (recipientEmails.length) {
    var inlineImages = draft.getAttachments({"includeAttachments": false, "includeInlineImages": true});
    var attachments = draft.getAttachments({"includeAttachments": true, "includeInlineImages": false});
    var plainBody = draft.getPlainBody().replace(RecipientsPattern, '').replace(SendDatePattern, '');
    var htmlBody = draft.getBody().replace(RecipientsPattern, '').replace(SendDatePattern, '');

    console.log("Sending email " + draft.getSubject() + " to " + recipientEmails);
    GmailApp.sendEmail(
        recipientEmails.join(","),
        draft.getSubject(), 
        replaceVariables(plainBody, recipient, recipientEmails, false),
        {
          "htmlBody": replaceVariables(htmlBody, recipient, recipientEmails, true),
          "attachments": attachments,
          "cc": draft.getCc(),
          "bcc": draft.getBcc(),
          "replyTo": "Marc Page <Marc@ResolveToExcel.com>",
          "name": "Marc Page",
          "inlineImages": matchAttachmentsFromBody(inlineImages, draft.getBody())
        });

    markRecipients(trackingSheet, recipientEmails);
  }
}

function afterDate(year, month, day) {
  var target = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
  var now = new Date();

  return now >= target;
}

function handleEmails() {
  var drafts = GmailApp.getDraftMessages();
  
  for (var draftIndex = 0; draftIndex < drafts.length; ++draftIndex) {
    var draft = drafts[draftIndex];
    var hasRecipients = draft.getBody().match(RecipientsPattern);
    var hasSendDate = draft.getBody().match(SendDatePattern);

    if (hasSendDate && !afterDate(hasSendDate[2], hasSendDate[3], hasSendDate[4])) {
      console.log("Waiting to send email " + draft.getSubject() + " until " + hasSendDate[0]);
      
      hasRecipients = false; // Not yet, wait
    }

    if (hasRecipients) {
      var all_contacts = getEveryone();
      var recipients = getEmailList(hasRecipients[1].split(','), all_contacts);
      var trackingSheet = getAutomatedEmailTracker(draft.getId(), draft.getSubject());

      for (var recipientIndex = 0; recipientIndex < recipients.length; ++recipientIndex) {
        sendMessage(draft, recipients[recipientIndex], trackingSheet['send_sheet'], trackingSheet['opt_out_sheet']);
      }
    }
  }
}
