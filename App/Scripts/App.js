/*
Checked in by: adambu
Date Created: 01/21/2018

Product Area Tags: Authentication / Authorization, Azure, Graph / OneDrive, 

Technology Tags: Client OM, JavaScript / JSOM, REST, 

Use Case: 
This is for uploading files larger than 4 MB to OneDrive or SharePoint Libraries using Graph API.

Description:
This sample uses ADAL.js in order to get an access token to the https://graph.microsoft.com resource.
It then lists the files in the OneDrive library (you can change to any SharePoint library if you can get the ID of the Drive in order to set the correct path),
and lets you upload a large file using the createUploadSession method.  There's a fair amount of logic you much create in order to get the "Content-Range" header
correctly formated for each PUT operation that much be called ot upload each chunk to the upload session that you create in the first POST to /createUploadSession.

To run the sample change the ClientId in the config.js file.

Keywords: Graph API ADAL.js OAuth2 JavaScript jQuery createUploadSession "Large Files" upload files 

*/

var main = function () {

    var baseEndpoint = "https://graph.microsoft.com";
    var filesResource = "https://" + tenant + "-my.sharepoint.com";
    var endpoints =  {
        filesResource:filesResource}; 

    // Enter Global Config Values & Instantiate ADAL AuthenticationContext
    window.config =  {
        tenant:tenant + ".onmicrosoft.com",
        clientId:clientId,
        postLogoutRedirectUri:window.location.origin,
        endpoints:endpoints,
        cacheLocation:"localStorage"// enable this for IE, as sessionStorage does not work for localhost.
    };
    var authContext = new AuthenticationContext(config); 

    var fileToUpload;

    // Get UI jQuery Objects
    var $filesPanel = $(".files-panel-body"); 
    var $userDisplay = $(".app-user"); 
    var $signInButton = $(".app-login"); 
    var $signOutButton = $(".app-logout"); 
    var $errorMessage = $(".app-error"); 
    var $fileControl = $("#trigger");
    var $btnUpload = $("#btnUpload");
    var $dataContainer = $(".data-container");
    var $dataLoadingCtl = $(".data-loading");

    // Check For & Handle Redirect From AAD After Login
    var isCallback = authContext.isCallback(window.location.hash); 
    authContext.handleWindowCallback(); 
    $errorMessage.html(authContext.getLoginError()); 

    if (isCallback &&  ! authContext.getLoginError()) {
        window.location = authContext._getItem(authContext.CONSTANTS.STORAGE.LOGIN_REQUEST); 
    }

    // Check Login Status, Update UI
    var user = authContext.getCachedUser(); 
    if (user) {
        $userDisplay.html(user.userName); 
        $userDisplay.show(); 
        $signInButton.hide(); 
        $signOutButton.show(); 
    }else {
        $userDisplay.empty(); 
        $userDisplay.hide(); 
        $signInButton.show(); 
        $signOutButton.hide(); 
    }

    // Register NavBar Click Handlers
    $signOutButton.click(function () {
        authContext.logOut(); 
    }); 
    $signInButton.click(function () {
        authContext.login(); 
    }); 

    $fileControl.change(function () {
        fileToUpload = event.srcElement.files[0];
    });

    $btnUpload.click(onUpload);

    refreshData();
    
    function getUploadSession(fileType, name) {
        console.log("getUploadSession method called::"); 

        const body =  {
            "item": {
                "@microsoft.graph.conflictBehavior":"rename"
            }
        }; 
        authContext.acquireToken(baseEndpoint, function (error, token) {

            // Handle ADAL Errors.
            if (error || !token) {
                printErrorMessage("ADAL error occurred: " + error);
                return;
            }

            //Execute POST request to get the UploadSession.
            $.ajax({
                type: "POST",
                /*
                    This uploads the chosen file to the users OneDrive library.  To Upload to 
                    a different library or folder, change the path here.  Use the Graph Explorer at:
                    https://developer.microsoft.com/en-us/graph/graph-explorer to help get the path or 
                    IDs or other document libraries.  
                */
                url: baseEndpoint + `/v1.0/me/drive/root:/${name}.${fileType}:/createUploadSession`,
                headers: {"Accept": "application/json, text/plain, */*", "Content-Type": "application/json", "Authorization": "Bearer " + token },
                body: body

            }).done(function (response) {
                console.log("Successfully got upload session.");
                console.log(response);

                var uploadUrl = response.uploadUrl;
                $dataLoadingCtl.html("Loading....");
                uploadChunks(fileToUpload, uploadUrl);

            }).fail(function (response) {
                console.log("Could not get upload session: " + response.responseText);
                
            });
        
        });
    }

    /*      After getting the uploadUrl, this function does the logic of chunking out 
            the fragments and sending the chunks to uploadChunk */
    async function uploadChunks(file, uploadUrl) {
        var reader = new FileReader(); 

        // Variables for byte stream position
        var position = 0; 
        var chunkLength = 320 * 1024; 
        console.log("File size is: " + file.size); 
        var continueRead = true; 
        while (continueRead) {
            var chunk; 
            try {
                continueRead = true; 
                //Try to read in the chunk
                try {
                    let stopByte = position + chunkLength; 
                    console.log("Sending Asynchronous request to read in chunk bytes from position " + position + " to end " + stopByte); 

                    chunk = await readFragmentAsync(file, position, stopByte); 
                    console.log("UploadChunks: Chunk read in of " + chunk.byteLength + " bytes."); 
                    if (chunk.byteLength > 0) {
                        continueRead = true; 
                    }else {
                        break; 
                    }
                    console.log("Chunk bytes received = " + chunk.byteLength); 
                }catch (e) {
                    console.log("Bytes Received from readFragmentAsync:: " + e); 
                    break; 
                }
                // Try to upload the chunk.
                try {
                    console.log("Request sent for uploadFragmentAsync"); 
                    let res = await uploadChunk(chunk, uploadUrl, position, file.size); 
                    // Check the response.
                    if (res.status !== 202 && res.status !== 201 && res.status !== 200)
                        throw ("Put operation did not return expected response"); 
                    if (res.status === 201 || res.status === 200)
                    {
                        console.log("Reached last chunk of file.  Status code is: " + res.status);
                        continueRead = false; 
                    }  
                    else
                    {
                        console.log("Continuing - Status Code is: " + res.status);
                        position = Number(res.json.nextExpectedRanges[0].split('-')[0]);
                    }    

                    console.log("Successful response received from uploadChunk."); 
                    console.log("Position is now " + position); 
                }catch (e) {
                    console.log("Error occured when calling uploadChunk::" + e); 
                }

            }catch (e) {
                continueRead = false; 
            }
        }    
        emptyDataContainer();
        refreshData();
    }

    // Reads in the chunk and returns a promise.
    function readFragmentAsync(file, startByte, stopByte) {
        var frag = ""; 
        const reader = new FileReader(); 
        console.log("startByte :" + startByte + " stopByte :" + stopByte); 
        var blob = file.slice(startByte, stopByte); 
        reader.readAsArrayBuffer(blob); 
        return new Promise((resolve, reject) =>  {
            reader.onloadend = (event) =>  {
                console.log("onloadend called  " + reader.result.byteLength); 
                if (reader.readyState === reader.DONE) {
                    frag = reader.result; 
                    resolve(frag); 
                }
            }; 
        }); 
    }

    // Upload each chunk using PUT
    function uploadChunk(chunk, uploadURL, position, totalLength) {
        var max = position + chunk.byteLength - 1; 
        //var contentLength = position + chunk.byteLength;

        console.log("Chunk size is: " + chunk.byteLength + " bytes."); 

        return new Promise((resolve, reject) =>  {
            console.log("uploadURL:: " + uploadURL); 

            try {
                console.log('Just before making the PUT call to uploadUrl.');
                let crHeader = "bytes " + position + "-" + max + "/" + totalLength;
                 //Execute PUT request to upload the content range.
                $.ajax({
                    type: "PUT",
                    contentType: "application/octet-stream",                    
                    url: uploadURL,
                    data: chunk,
                    processData: false,
                    headers: {"Content-Range": crHeader }                    

                }).done(function (data, textStatus, jqXHR) {
                    console.log("Content-Range header being set is : " + crHeader);
                    if (jqXHR.responseJSON.nextExpectedRanges) {
                        console.log("Next Expected Range is: " + jqXHR.responseJSON.nextExpectedRanges[0]);
                    }
                    else {
                        console.log("We've reached the end of the chunks.")
                    }                        
                    
                    results = {};
                    results.status = jqXHR.status;
                    results.json = jqXHR.responseJSON;
                    resolve(results);

                }).fail(function (response) {
                    console.log("Could not upload chunk: " + response.responseText);    
                    console.log("Content-Range header being set is : " + crHeader);  

                });
            }catch (e) {
                console.log("exception inside uploadChunk::  " + e); 
                reject(e); 
            }
        }); 
    }

    function clearErrorMessage() {
        var $errorMessage = $(".app-error"); 
        $errorMessage.empty(); 
    }

    function printErrorMessage(mes) {
        var $errorMessage = $(".app-error"); 
        $errorMessage.html(mes); 
    }

    function onUpload() {
        var uploadUrl; 
        try {
            var i = fileToUpload.name.lastIndexOf("."); 
            clearErrorMessage();
        }
        catch (error) {
            printErrorMessage ("Please select a file to upload first.")
            return;
        }

        var fileType = fileToUpload.name.substring(i + 1); 
        var fileName = fileToUpload.name.substring(0, i); 
        getUploadSession(fileType, fileName);  

    }

    function refreshData() {

        console.log("Fetching files using Graph API");

        // Acquire token for Files resource.
        authContext.acquireToken(baseEndpoint, function (error, token) {

            // Handle ADAL Errors.
            if (error || !token) {
                printErrorMessage("ADAL error occurred: " + error);
                return;
            }

            //Execute GET request to Files API.
            //Refer to the API reference for more information: https://msdn.microsoft.com/en-us/office/office365/api/files-rest-operations
            $.ajax({
                type: "GET",
                //url: baseEndpoint + "/_api/v1.0/me/files",
                url: baseEndpoint + "/v1.0/me/drive/root/children",
                headers: { "Authorization": "Bearer " + token }

            }).done(function (response) {
                console.log("Successfully fetched files from OneDrive.");
                console.log(response);

                var $html = $(".panel-body");
                var $template = $html.find(".data-container");
                var output = "";

                response.value.forEach(function (item) {
                    var $entry = $template;
                    var typeVal = item["folder"] ? "Folder" : "File";
                    $entry.find(".view-data-type").html(typeVal);
                    $entry.find(".view-data-name").html(item.name);
                    output += $entry.html();
                });

                // Update the UI.
                $dataContainer.html(output);

            }).fail(function () {
                console.log("Fetching files from OneDrive failed.");
                printErrorMessage("Something went wrong! Try refreshing the page.");
            });
        });
    }

    function emptyDataContainer() {
        var dataRows = $dataContainer.find(".data-row");

        while (dataRows.length > 1)
        {
            dataRows[dataRows.length - 1].remove();
            dataRows = $dataContainer.find(".data-row");
        }            
        dataRows.find(".view-data-type").empty();
        dataRows.find(".view-data-name").empty();
    }
}

$(document).ready(function(){
   main();
});