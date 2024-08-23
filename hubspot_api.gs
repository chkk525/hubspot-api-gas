function hubspotApi(apiEndPoint, method, version = 3, payload = null) {
  let url = `https://api.hubapi.com/crm/v${version}/${apiEndPoint}`
  const accessToken = PropertiesService.getScriptProperties().getProperty('hs_access_token')
  const headers = {
    'Authorization': 'Bearer ' + accessToken
  };
  const options = {
    method: method,
    headers: headers,
    contentType: "application/json",
    muteHttpExceptions: true,
  };
  if (payload !== null) {
    payload = JSON.stringify(payload)˛
    options["payload"] = payload
  }


  let response, parsedResponse, results = []
  while (true) {
    response = UrlFetchApp.fetch(url, options)
    if (response.getResponseCode() === 200) {
      parsedResponse = JSON.parse(response.getContentText())
      results.push(...parsedResponse['results'])

      
      if (parsedResponse['paging'] && parsedResponse['paging']['next']) {
        url = parsedResponse['paging']['next']['link']
      } else {
        break
      }
    }
  }

  return results;
}

function parseDate(txt) {
  const d = new Date(txt)
  if (txt == null | isNaN(d.getMonth())) {
    return ''
  }
  else {
    return d
  }
}

function dataToSheet(data, sheet_name) {
  // Spreadsheet ID
  const spreadsheet = SpreadsheetApp.openById('1pk83VeTJ2UoWar0ZCV5yvL9tSfPsMyrxby3lLEfbNNM')
  let sheet = spreadsheet.getSheetByName(sheet_name)

  if (!sheet) {
    sheet = spreadsheet.insertSheet()
    sheet.setName(sheet_name)
  }

  const nrow = data.length
  const ncol = data[0].length
  sheet.clear()
  sheet.getRange(1, 1, nrow, ncol).setValues(data)
}

function getAllProperties(objectType) {
  const results = hubspotApi(`properties/${objectType}`, 'get')
  const dataHeaders = ['name', 'label', 'description', 'groupName']
  const properties = [dataHeaders]
  results.forEach(function (result) {
    row = dataHeaders.map(col => result[col])
    properties.push(row)
  })
  dataToSheet(properties, `_${objectType}_properties`)
}

function getAllDealsProperties() {
  getAllProperties('deal')
}

function getAllContactsProperties() {
  getAllProperties('contact')
}

function getContacts() {
  const properties = [
    "firstname",
    "lastname",
    "email",
    "company",
    "hs_analytics_source_data_1",
    "hs_analytics_source_data_2",
    "hs_latest_source_data_1",
    "hs_latest_source_data_2",
    "hs_email_last_email_name",
    "hs_lead_status",
    "ip_city",
    "hs_analytics_num_page_views",
    "listing",
    "hs_sa_first_engagement_date",
    "hs_sa_first_engagement_descr",
    "hs_analytics_first_touch_converting_campaign",
    "hs_analytics_last_touch_converting_campaign",
    "first_conversion_date",
    "first_conversion_event_name",
    "first_deal_created_date",
    "from_where",
    "notes_last_updated",
    "hs_analytics_first_url",
    "rsormpg",
    "first_deal_created_date",
    "moniplat______",
    "leadsource",
    "moniplat_lead_created_at",
    "ap__",
    "ap___",
    "ap____",
    "ap_____",
    "ap___id",
    "ap_introtime",
    "apcompany_id",
    "apdl__",
    "apposition",
    "exclude",
    "excludereason",
    "sendtime"
  ]

  const datetimeColumns = ["first_conversion_date", "notes_last_updated","first_deal_created_date","hs_sa_first_engagement_date"]

  const propertiesQuerys = properties.map(prop => `properties=${prop}`).reduce((prev, current) => `${prev}&${current}`)
  const results = hubspotApi(`objects/contacts?${propertiesQuerys}&limit=100`, 'get', version = 3)

  const contacts = [['id', ...properties]]
  results.forEach(function (result) {
    const id = result['id']
    const propertyValues = properties.map(prop => {
      if (datetimeColumns.includes(prop)) {
        return parseDate(result['properties'][prop])
      }
      return result['properties'][prop]
    })
    contacts.push([id, ...propertyValues])
  })

  dataToSheet(contacts, '_contacts')
}

function getDeals() {
  const properties = [
    "hs_object_id",
    "hs_analytics_source_data_1",
    "hs_analytics_source_data_2",
    "notes_last_updated",
    "dealstage",
    "pipeline",
    "amount",
    "dealname",
    "division",
    "hubspot_owner_id"
  ]

  const datetimeColumns = ["first_conversion_date", "notes_last_updated"]

  const propertiesQuerys = properties.map(prop => `properties=${prop}`).reduce((prev, current) => `${prev}&${current}`)
  const results = hubspotApi(`objects/deals?${propertiesQuerys}&limit=100&associations=contacts`, 'get', version = 3)

  const contacts = [[...properties, "contact_id0", "contact_id1", "contact_id2"]]
  results.forEach(function (result) {
    const propertyValues = properties.map(prop => {
      if (datetimeColumns.includes(prop)) {
        return parseDate(result['properties'][prop])
      }
      return result['properties'][prop]
    })

    let contactIds = []
    if ('associations' in result) {
       contactIds = result['associations']['contacts']['results'].map(el=>el['id'])
    }
    contactIds = comlementArrayLength(contactIds, 3)
    contacts.push([...propertyValues, ...contactIds])
  })
  dataToSheet(contacts, '_deals')
}

function comlementArrayLength(arr,len){  
  if(arr.length < len){
    const compArr = Array(len - arr.length).fill('')
    arr = [...arr, ...compArr]
  } else {
    arr = arr.slice(0,len)
  }
  return arr
}

function refresh(){
  getContacts()
  getDeals()
}

