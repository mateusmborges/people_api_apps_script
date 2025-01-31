function importAllContacts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  sheet.clear();

  // Table Header
  let header = [['Name', 'Email', 'Phone', 'Birthday', 'Company', 'Position', 'Obs.', 'Labels']];

  let fullList = [];

  let contacts = People.People.Connections.list("people/me",{
      personFields: 'names,emailAddresses,phoneNumbers,birthdays,occupations,organizations,biographies,memberships',
      pageSize: 1000
    });

  let dict = getLabels(contacts);

  do {
    contacts = People.People.Connections.list("people/me",{
      personFields: 'names,emailAddresses,phoneNumbers,birthdays,occupations,organizations,biographies,memberships',
      pageSize: 1000,
      pageToken: contacts.nextPageToken
    });

    contactList = contacts['connections'].map(function(connection) {
      const name = connection['names'] ? connection['names'][0]['displayName'] : "No Name";
      const email = connection['emailAddresses'] ? connection['emailAddresses'][0]['value'] : "No Email";
      const phone = connection['phoneNumbers'] ? connection['phoneNumbers'][0]['value'] : "No Phone";
      const birthday = connection['birthdays'] ? parseDate(connection['birthdays'][0]['date']) : ""; 
      const company = connection['organizations'] ? connection['organizations'][0]['name'] : "";
      const position = connection['organizations'] ? connection['organizations'][0]['title'] : "";
      const obs = connection['biographies'] ? connection['biographies'][0]['value'] : "";
      const labels = connection['memberships'] ? connection['memberships']
        .filter(membership => membership['contactGroupMembership']['contactGroupId'] != "myContacts")
        .map(membership => dict[membership['contactGroupMembership']['contactGroupResourceName']])
        .join(",") : "";

      return [name, email, phone, birthday, company, position, obs, labels];
    });

    fullList = fullList.concat(contactList);
    //fullList.sort();

  } while (contacts.nextPageToken);

  let data = header.concat(fullList);

  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
}

function parseDate(date) {
  const dia = date.day ? (date.day >= 10 ? date.day : '0' + date.day.toString()) : "--";
  const mes = date.month ? (date.month >= 10 ? date.month : '0' + date.month.toString()) : "--";
  const ano = date.year ? date.year : "----";
  
  return `${dia}/${mes}/${ano}`;
}

function getLabels(contacts) { 
  let resourceNames = [];
  
  contacts['connections'].forEach(function(connection) {
    if(connection['memberships']) {
      connection['memberships'].forEach(function(membership) {
        let resourceName = membership['contactGroupMembership']['contactGroupResourceName'];
        resourceNames.push(resourceName);
      });
    }
  });

  resourceNames = [...new Set(resourceNames)];

  let dictNames = {};

  resourceNames.forEach(function(resourceName) {
    dictNames[resourceName] = People.ContactGroups.get(resourceName)['name'];
  });

  return dictNames;  
}
