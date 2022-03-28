<cfset variables.status     = structNew() />
<cfset variables.userObj    = createObject("component", "local.cfc.userdetails") />
<cfset variables.hasExcel   = false />
<cfset variables.queryResultSet = '' />
<cfset variables.downloadExcel  = '' />
<cfif cgi.request_method EQ "post" AND form.formtype EQ 'submit_excel'>
    <cfparam  name="form.excelFile" default="">
    <cfif len(trim(form.excelFile))>
            <cftry>
                    <cffile action            = "upload" 
                        fileField             = "excelFile" 
                        destination           = "C:\ColdFusion2021\cfusion\wwwroot\userdetails\uploads\" 
                        allowedExtensions     = ".xlsx,.xls"
                        nameConflict          = "overwrite"
                        result                =  variables.excelData  
                    > 
                    <cfset variables.uploadedPath = "#variables.excelData.serverdirectory#\#variables.excelData.serverfile#"/>
                    <cfspreadsheet 
                            action           = "read" 
                            src              = "#variables.uploadedPath#"
                            headerrow        = 1
                            excludeHeaderRow = true
                            query            = "queryData" 
                    />
                    <cfset variables.processExcel  =  variables.userObj.processMyExcel(queryData) />
                    <cfif isQuery(variables.processExcel) EQ false AND structKeyExists(variables, "processExcel") AND variables.processExcel EQ "empty_excel">
                        <cfset variables.status.data    = 'error' />
                        <cfset variables.status.message = 'No data to process. You have uploaded an empty excel file.' />
                    <cfelse>
                        <cfset variables.status.data    =   'success' />
                        <cfset variables.status.message =   'Your excel has been uploaded successfully.' />
                        <cfif isQuery(variables.processExcel) EQ true>
                            <cfset variables.hasExcel           =   true />
                            <cfset variables.queryResultSet     =   variables.processExcel />
                        </cfif>
                    </cfif>
                <cfcatch type="exception">
                    <cfset variables.status.data    = 'error' />
                    <cfset variables.status.message = 'Exception: #cfcatch.message#' />
                </cfcatch>
            </cftry>
        <cfelse>
            <cfset variables.status.data    = 'error' />
            <cfset variables.status.message = 'Please upload a valid excel file' />
    </cfif>
</cfif>
<cfif cgi.request_method EQ "post" AND form.formtype EQ 'submit_download'>
    <cfset variables.jsonText           =   form.querytext />
    <cfset variables.downloadExcel      =   variables.userObj.downloadResults(variables.jsonText) />
</cfif>
<!doctype html>
<html lang="en">
    <head>
        <meta charset="utf-8">
        <title>User Details</title>
        <link href="dist/css/bootstrap.min.css" rel="stylesheet">
        <link href="dist/css/styles.css" rel="stylesheet">
    </head>
    <body>
        <cfoutput>
            <main class="container">
                <h1 class="text-center mt-3">User Information</h1>
                <div class="px-4 py-5 my-5 text-center">
                    <div class="row">
                        <div class="col-lg-4">
                            <a href="./cfc/userdetails.cfc?method=downloadData&type=plain" class="btn btn-sm btn-success">Plain Template</a>
                            <a href="./cfc/userdetails.cfc?method=downloadData&type=data" class="btn btn-sm btn-info">Template With Data</a>
                        </div>
                        <div class="col-lg-4">
                            <cfif variables.hasExcel    ==   true>
                                <cfset variables.querytojson    =  SerializeJSON(variables.queryResultSet) />
                                <form action="" method="post">
                                    <input type="hidden" name="formtype" value="submit_download" />
                                    <textarea name="querytext" style="visibility:hidden;">#variables.querytojson#</textarea>
                                    <button type="submit" class="btn btn-sm btn-success mb-3 col-lg-6">Download</button>
                                </form>
                            </cfif>
                        </div>
                        <div class="col-lg-4">
                            <form action="" method="post" enctype="multipart/form-data">
                                <div class="row">
                                    <div class="col-lg-6">
                                        <input class="form-control form-control-sm" id="formFileSm" type="file" name="excelFile">
                                    </div>
                                    <input type="hidden" name="formtype" value="submit_excel" />
                                    <button type="submit" class="btn btn-sm btn-success mb-3 col-lg-6">Upload</button>
                                </div>
                            </form>
                        </div>
                    </div>
                    <div class="col-lg-8 mx-auto">
                            <cfif structKeyExists(status, 'data')>
                                <cfif variables.status.data EQ 'success'>
                                    <div class="alert alert-primary" role="alert">
                                        #variables.status.message#
                                    </div>
                                <cfelseif variables.status.data EQ 'error'>
                                    <div class="alert alert-danger" role="alert">
                                        #variables.status.message#
                                    </div>
                                </cfif>   
                            </cfif>
                        <table class="table">
                            <thead>
                            <tr>
                                <th scope="col"></th>
                                <th scope="col">First Name</th>
                                <th scope="col">Last Name</th>
                                <th scope="col">Address</th>
                                <th scope="col">Email</th>
                                <th scope="col">Phone</th>
                                <th scope="col">DOB</th>
                                <th scope="col">Role</th>
                            </tr>
                            </thead>
                            <tbody>
                                <cfset userData  = userObj.getUsers() />
                                <cfloop query="userData">
                                    <tr>
                                        <th scope="row">#userData.id#</th>
                                        <td>#userData.first_name#</td>
                                        <td>#userData.last_name#</td>
                                        <td>#userData.address#</td>
                                        <td>#userData.email#</td>
                                        <td>#userData.phone#</td>
                                        <td>#userData.dob#</td>
                                        <td>#userData.roles#</td>
                                        <td></td>
                                    </tr>
                                </cfloop>
                            </tbody>
                        </table>
                    </div>
                </div>
            </main>
        </cfoutput>
    <script src="dist/js/bootstrap.bundle.min.js"></script>
    </body>
</html>
