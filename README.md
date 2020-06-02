Sharepoint using .NET CORE 3.1 and Microsoft SharePoint Client component library

Description:

Endpoints:
* Get all documents in library (This endpoint does not support 5000+ entities).
    GET /api/sharepoint/documents?site=<sitename>&list=<listname>
* List of libraries from a SharePoint site
    GET /api/sharepoint/lists?site=<sitename>
* List of folders/documentsets from a sharepoint library
    GET /api/sharepoint/folders?site=<sitename>&list=<listname>
* documents with metadata from sharepoint library
    GET /api/sharepoint/documentswithfields?site=<sitename>&list=<listname>
* List of available fields on specific library
    GET /api/sharepoint/fields?site=<sitename>&list=<listname>
* Return user id
    GET /api/sharepoint/userid?name=<email>
* Array of folder names
    GET /api/sharepoint/foldernames?site=<sitename>&list=<listname>
* Create new document
    POST /api/sharepoint/newdocument
* Delete a site
    DELETE /api/sharepoint/deletesite
    {
       "site": <"site name">
    }
* Create documentset
    POST /api/sharepoint/documentset
    {
        "site": <"site name">,
        "list" :<"list name">,
        "sitecontent" : <"site content name">,
        "documentset" : <"name of the new document set">,
     } 
* SystemUpdate metadata
    POST /api/sharepoint/updatemetadata
    {
    	"ListName":"Documents",
    	"FileName":"Cyan.svg",
    	"FolderName":"My first document set",
    	"Fields":{
    			"BLAD":"9",
    			"BESKRIVELSE":"Beskrivelse updated",
    			"DOC_NO": "123433334455",
    			"DATO":"2020-01-01 04:00:00"
    
    	}
    }  
* Upload file to sharepoint
    POST /api/sharepoint/UploadToSharePoint
    {
        "list":"Dokumentasjon",
        "file_url":"http://.....",
        "foldername":"Landskaps og miljøplan",
        "site": "sporaevk",
        "filename": "Postnummerregister-Excel.xlsx"
    }
* Upload file through SMB fileshare and update metadata
    POST /api/sharepoint/migration
* Upload file through SMB fileshare and update metadata
    POST /api/sharepoint/migrationoptimize
* Update metadata.
    POST /api/sharepoint/documentfix
    NB! works only on library that has eDocsDokumentnavn field name. Use only on lists with over 5000 documents
* Update existing document with SystemUpdate() to prevent version increment.
    POST /api/sharepoint/document
* Enrich metadata on documentset only
    POST /api/document/folderenrichment
* Enrich sharepoint library with overwriting version on library with 5000+ documents
    POST /api/document/updateoverwriteversion
* Enrich sharepoint library with 5000+ documents
    POST /api/document/UpdateWithoutVersioning
* Migration on library with versioning (used for existing library with 5000+)
    POST /api/document/MigrationWithVersioning
* List of documents (5000+)
    GET /api/document/all





