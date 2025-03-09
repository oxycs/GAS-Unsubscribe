// code breaks once we hit 1000 contacts, fyi, though we're nowhere near that

const formResponseSheetID = "";

function mailListResourceName() {
  try {
    // get all contact groups (labels)
    const clist = People.ContactGroups.list();
    // for each group...
    for (l=0;l<clist.contactGroups.length;l++) {
      // if the name of the group is "mailing list"...
      if (clist.contactGroups[l].formattedName == "Mailing List") {
        // return the reosurce name (string, looks like "contactGroup/********")
        return clist.contactGroups[l].resourceName;
      }
    }
  } catch (err) {
    console.log('Failed to get the connection with an error %s', err.message);
  }
}

function mListObj() {
  try {
    // get mailing list object, including up to 1000 members of list
    const mlist = People.ContactGroups.get(mailListResourceName(), {
      maxMembers: 1000
    });
    // return object
    return mlist;
  } catch (err) {
    console.log('Failed to get the connection with an error %s', err.message);
  }
}

function unsub(toUnsub) {
  // let toUnsub = "gilbertadler@oxy.edu";
  try {
    // Get all contacts (up to 1000)
    const people = People.People.Connections.list('people/me', {
      personFields: 'emailAddresses',
      pageSize: 1000
    });
    // for each contact
    for (i = 0; i < people.connections.length; i++) {
      // for each email in each contact
      for (j = 0; j < people.connections[i].emailAddresses.length; j++) {
        // if the email in the contact is the one that we want to unsubscribe...
        if (people.connections[i].emailAddresses[j].value === toUnsub) {
          // get the mailing list object
          let mlist = mListObj();
          // for each member of the mailing list
          for (k=0; k<mlist.memberResourceNames.length; k++) {
            // if the person we want to unsubscribe is the selected member of the mailing list...
            if (people.connections[i].resourceName == mlist.memberResourceNames[k]) {
              // remove user from the mailing list
              People.ContactGroups.Members.modify({ resourceNamesToRemove: [people.connections[i].resourceName] }, mailListResourceName())
            }
          }
          break;
        }
      }
    }
  } catch (err) {
    console.log('Failed to get the connection with an error %s', err.message);
  }
}

function main() {
  // get sheet object
  let ss = SpreadsheetApp.openById(formResponseSheetID).getSheets()[0]
  // get last response row
  let lastRow = ss.getLastRow();
  // get all responses from first to last, as 2D array
  let range = ss.getRange(`A2:D${lastRow}`)
  let content = range.getValues()
  for (m=0; m<content.length; m++) {
    // if attempted not true...
    if (content[m][2] != true) {
      // set attempted to true
      content[m][2] = true;
      range.setValues(content);
      Utilities.sleep(1000)
      try {
        // attempt to unsub
        unsub(content[m][1])
        // set success to true
        content[m][3] = true;
      } catch {
        console.log("FAILED TO UNSUBSCRIBE");
      }
      range.setValues(content);
      Utilities.sleep(1000)
    }
  }
}
