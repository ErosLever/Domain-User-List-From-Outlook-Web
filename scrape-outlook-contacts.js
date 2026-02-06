// Download text to a file
// https://stackoverflow.com/a/47359002
function saveAs(text, filename) {
    var pom = document.createElement('a');
    pom.setAttribute('href', 'data:text/plain;charset=urf-8,' + encodeURIComponent(text));
    pom.setAttribute('download', filename);
    pom.click();
};

// Convert 2d array into comma separated values
// https://stackoverflow.com/a/14966131
function convertToCsv(rows) {
    //let csvContent = "data:text/csv;charset=utf-8,";
    const s = ',';    // separator
    const q = '"';    // quote
    const l = '\r\n'; // line ending
    let csvContent = ["ADObjectId","EmailAddress","DisplayName"].join(s) + l;

    rows.forEach(function(rowArray) {
        let row = rowArray.map(x => x.includes(q) || x.includes(s) ? `${q}${x.replaceAll(q,q+q)}${q}` : x ).join(s);
        csvContent += row + l;
    });
    return csvContent;
}

var x_owa_urlpostdata = {
    "__type": "FindPeopleJsonRequest:#Exchange",
    "Header": {
        "__type": "JsonRequestHeaders:#Exchange",
        "RequestServerVersion": "V2018_01_08",
        "TimeZoneContext": {
            "__type": "TimeZoneContext:#Exchange",
            "TimeZoneDefinition": {
                "__type": "TimeZoneDefinitionType:#Exchange",
                "Id": "W. Europe Standard Time"
            }
        }
    },
    "Body": {
        "__type": "FindPeopleRequest:#Exchange",
        "IndexedPageItemView": {
            "__type": "IndexedPageView:#Exchange",
            "BasePoint": "Beginning",
            "Offset": 1337.1337,
            "MaxEntriesReturned": 150
        },
        "PersonaShape": {
            "__type": "PersonaResponseShape:#Exchange",
            "BaseShape": "Default",
            "AdditionalProperties": [{
                "__type": "PropertyUri:#Exchange",
                "FieldURI": "PersonaAttributions"
            }, {
                "__type": "PropertyUri:#Exchange",
                "FieldURI": "PersonaRelevanceScore"
            }, {
                "__type": "PropertyUri:#Exchange",
                "FieldURI": "PersonaTitle"
            }]
        },
        "ShouldResolveOneOffEmailAddress": false,
        "QueryString": "PLACEHOLDER",
        "SearchPeopleSuggestionIndex": true,
        "Context": [{
            "__type": "ContextProperty:#Exchange",
            "Key": "AppName",
            "Value": "OWA"
        }, {
            "__type": "ContextProperty:#Exchange",
            "Key": "AppScenario",
            "Value": "peopleHubReact"
        }, {
            "__type": "ContextProperty:#Exchange",
            "Key": "DisableAdBasedPersonaIdForPersonalContacts",
            "Value": "true"
        }],
        "QuerySources": [
            "Directory"
        ]
    }
};

// Get Bearer Token from LocalStorage
var keys = Object.keys(localStorage).filter(
    x => x.startsWith('msal') && /login\.windows\.net/.test(x) && /accesstoken/.test(x) && /outlook\.office\.com/.test(x)
);

var local_storage_stuff = JSON.parse(localStorage[keys[0]]);
var bearer_token = local_storage_stuff.secret;
var tenant_id = local_storage_stuff.realm;





async function fetchAllItems() {
    let user_db = new Set();

    // Listing all users requires Directory.Read.All which is hardly the case for unprivileged users
    // instead we perform a basic search of single characters
    var queries = [].map.call("1234567890qwertyuiopasdfghjklzxcvbnm", x => x);
    while (queries.length > 0) {

        let query = queries.shift();

        // Outlook forces max 150 results pagination, we keep track of the offset
        let offset = 0;
        while (true) {

            // Compose actual URL with offset and search query (single character)
            let x_owa_urlpostdata_offset = JSON.stringify(x_owa_urlpostdata)
                .replace('1337.1337', offset)
                .replace('PLACEHOLDER', query);

            const response = await fetch("https://outlook.office.com/owa/service.svc?action=FindPeople&app=People", {
                "headers": {
                    "accept": "*/*",
                    "accept-language": "en-US;q=0.9,en;q=0.8",
                    "action": "FindPeople",
                    "authorization": "Bearer " + bearer_token,
                    "cache-control": "no-cache",
                    "content-type": "application/json; charset=utf-8",
                    "pragma": "no-cache",
                    "prefer": "IdType=\"ImmutableId\", exchange.behavior=\"IncludeThirdPartyOnlineMeetingProviders\"",
                    "x-owa-urlpostdata": encodeURIComponent(x_owa_urlpostdata_offset),
                    "x-req-source": "People",
                    "x-tenantid": tenant_id
                },
                "body": null,
                "method": "POST",
                "mode": "cors",
                "credentials": "include"
            });

            if (!response.ok) {
                throw new Error(`HTTP ${response.status}`);
            }

            const data = await response.json();
            let users = data.Body.ResultSet;

            if (users.length === 0)
                break

            users.map(user => [
                user.ADObjectId,
                user.EmailAddress.EmailAddress,
                user.DisplayName
            ]).map(x => user_db.add(x));

            offset += users.length;
        }
    }

    let export_filename = `office365_export_${tenant_id}_${new Date().toISOString().replace(/\D/g,'')}.csv`

    // Download database as a .csv file
    console.debug('Converting user array to csv...')
    let user_db_csv = convertToCsv(user_db);
    console.debug('Downloading csv...')
    saveAs(user_db_csv, export_filename);
    console.log('Downloaded results to "' + export_filename + '"!');
    return user_db;
}

fetchAllItems()
    .then(items => console.log("Fetched:", items.length))
    .catch(console.error);