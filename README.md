# Office365 User Scraper
Export everyone in your Office 365 organisation into a `.csv` in seconds, straight from your browser, without any admin tools.

This project reverse engineers the Office 365 Outlook webapp API, collecting all users in an Outlook directory via ~~a single~~ API requests. All you have to do is ~~obtain a `BaseFolderID` and~~ paste JavaScript code into your browser console. All users within that ~~BaseFolder~~ Directory will be downloaded to a .csv file on your computer.

The `x-owa-canary` cookie is ~~automatically retrieved from your browser and used to authenticate the API request~~ no longer required. Authentication is performed retrieving the Authorization Bearer token from LocalStorage.

The API response is then parsed and entered into a 2d array. This array is then converted into a comma-separated-value format which is then downloaded as a `.csv` file via your browser.

![Screenshot of the console output](https://smcclennon-img.netlify.app/projects/ous/console-output.png)

### Sample csv
Below is the data structure of the `.csv` file generated, formatted as a Markdown table (information redacted):
|ADObjectId|EmailAddress|DisplayName|
|:-:|:-:|:-:|
|2f9cef48-d156-4508-b559-fe947717aeee|P___@domain.org|Aa___ P___|
|4304d3d2-b258-47a2-bedc-32498cf4fb07|E___@domain.org|Ab___ E___|
|37a612c1-36b5-48f8-806d-869dea9dc6c4|G___@domain.org|Ab___ G___|
|d422f790-7d55-4655-866d-5f96a74fe29f|C___@domain.org|Ac___ C___|
|303d1d64-e3b9-4dd6-811a-b029c65c7972|a___@domain.org|Ad___ R___|
|303b10c9-a559-4124-9abc-b8884bfd3337|a___@domain.org|AF___ R___|
|6f466f00-3408-43aa-8b4c-5ca3f200604f|N___@domain.org|Ai___ N___|
|ea468f94-8937-4071-b3a2-e8fe468be551|D___@domain.org|Al___ D___|

## Features
- Retrieve full name, email address and Active Directory ObjectId by default
- Export all users to a `.csv` file
- Very portable, simply paste code into your browser console
- ~~Quiet network traffic (only 1 request)~~
- Automatically retrieve the required cookie for API authentication
- Easily extract more information from the API response (see [appendix](#Example-API-response))

## Usage
1. Visit https://outlook.office.com in your browser.
2. Press `F12` to launch the "Developer Tools" popup.
3. Navigate to the "Console" tab within the Developer Tools popup.
4. Paste [this](https://raw.githubusercontent.com/ErosLever/Domain-User-List-From-Outlook-Web/master/scrape-outlook-contacts.js) JavaScript code into the Developer Tools Console.
5. Press `Enter` to execute the code in the Console. Userdata will be ~~printed to the console and~~ downloaded to your computer shortly. If something went wrong, you will receive a JavaScript error in the Console, so make sure your Console is not filtering out errors.

## Appendix

### Example API response
The API responds with a list of users. Below is the data structure returned per user (some information redacted):
```json
[
    {
      "__type": "PersonaType:#Exchange",
      "PersonaId": {
        "__type": "ItemId:#Exchange",
        "Id": "AAUQAGAxxxxxxxxxxxxxxxxT2rY="
      },
      "PersonaTypeString": "Person",
      "CreationTimeString": "0001-01-02T00:00:00Z",
      "DisplayName": "Aa___ P___",
      "DisplayNameFirstLast": "Aa___ P___",
      "DisplayNameLastFirst": "Aa___ P___",
      "FileAs": "",
      "GivenName": "Aa___",
      "Surname": "P___",
      "CompanyName": "My Domain",
      "EmailAddress": {
        "Name": "Aa___ P___",
        "EmailAddress": "P___@domain.org",
        "RoutingType": "SMTP",
        "MailboxType": "Mailbox"
      },
      "EmailAddresses": [
        {
          "Name": "Aa___ P___",
          "EmailAddress": "P___@domain.org",
          "RoutingType": "SMTP",
          "MailboxType": "Mailbox"
        }
      ],
      "ImAddress": "sip:p___@domain.org",
      "WorkCity": "Watford",
      "RelevanceScore": 2147483647,
      "AttributionsArray": [
        {
          "Id": "0",
          "SourceId": {
            "__type": "ItemId:#Exchange",
            "Id": "AAUQAGAxxxxxxxxxxxxxxxxT2rY="
          },
          "DisplayName": "GAL",
          "IsWritable": false,
          "IsQuickContact": false,
          "IsHidden": false,
          "FolderId": null,
          "FolderName": null,
          "IsGuest": false
        }
      ],
      "ADObjectId": "aa000000-a0aa-00a0-0000-aaa000a0aaa0"
    }
]
```

### x-owa-urlpostdata decoded
`x-owa-urlpostdata` is a header used in the POST request to the Outlook API. We customise the following values in this header:
- `Offset`: Starting index of users to send. An offset of `20` will not return the first 20 users. By default, an offset of `0` is used to return all users.
- `MaxEntriesReturned`: Maximum number of users to be returned by the API. See the [Example API response](#Example-API-response) appendix to view the information returned per user. ~~By default, we request a maximum of `1000` users to be returned. However, you can increase this if you need to~~. This is enforced to maximum 150, requires paging using the `offset` parameter.
- `BaseFolderId Id`: This is the Outlook userlist/directory to return in the API response. This is tedious to obtain, but is essential and must be valid or the API request will fail. Note: not necessary if we are limiting the search to the `Directory` only as `QuerySources`.
```json
{
    "__type": "FindPeopleJsonRequest:#Exchange",
    "Header": {
        "__type": "JsonRequestHeaders:#Exchange",
        "RequestServerVersion": "V2018_01_08",
        "TimeZoneContext": {
        "__type": "TimeZoneContext:#Exchange",
        "TimeZoneDefinition": {
            "__type": "TimeZoneDefinitionType:#Exchange",
            "Id": "GMT Standard Time"
        }
        }
    },
    "Body": {
        "IndexedPageItemView": {
        "__type": "IndexedPageView:#Exchange",
        "BasePoint": "Beginning",
        "Offset": Offset,
        "MaxEntriesReturned": MaxEntriesReturned
        },
        "QueryString": null,
        "ParentFolderId": {
        "__type": "TargetFolderId:#Exchange",
        "BaseFolderId": {
            "__type": "AddressListId:#Exchange",
            "Id": BaseFolderId
        }
        },
        "PersonaShape": {
        "__type": "PersonaResponseShape:#Exchange",
        "BaseShape": "Default",
        "AdditionalProperties": [
            {
            "__type": "PropertyUri:#Exchange",
            "FieldURI": "PersonaAttributions"
            },
            {
            "__type": "PropertyUri:#Exchange",
            "FieldURI": "PersonaTitle"
            },
            {
            "__type": "PropertyUri:#Exchange",
            "FieldURI": "PersonaOfficeLocations"
            }
        ]
        },
        "ShouldResolveOneOffEmailAddress": false,
        "SearchPeopleSuggestionIndex": false
    }
}
```

## Special thanks
- [`@smcclennon`](https://github.com/smcclennon) - Original author of the [Outlook User Scaper](https://github.com/smcclennon/ous) project.
- [`@edubey`](https://github.com/edubey) - [Inspiration](https://github.com/edubey/browser-console-crawl/blob/master/single-story.js) for the project.
- [`@freddierick`](https://github.com/freddierick) - [Async fix](https://github.com/smcclennon/ous/commit/7ae0bc62468ddddc435481b7dae3abad8800890c) to make everything wait for the API `fetch()` request to complete.
