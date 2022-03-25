component displayname="userdetails"
    {
        remote function downloadData(required string type)
            {
                if(arguments.type == 'plain')
                    {
                        objSpreadsheet          = SpreadsheetNew("Sheet1",true);
                        SpreadsheetAddRow( objSpreadsheet, "First Name, Last Name, Address, Email, Phone, DOB, Role" );
                        SpreadsheetFormatRow( objSpreadsheet, {bold=true, alignment="center"}, 1 );
                        cfheader(
                                    name="Content-Disposition",
                                    value="attachment; filename=Plain_Template.xlsx"
                                );
                        cfcontent(
                                    type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    variable="#SpreadsheetReadBinary( objSpreadsheet )#"
                                );
                    }
                else if(arguments.type == 'data')
                    {
                        local.userSet           =   getUsers();
                        local.excelOutputQuery  =   queryNew("FirstName, LastName, Address, Email, Phone, DOB, Role, Result");
                        for(local.row IN local.userSet)
                            {
                                queryAddRow(local.excelOutputQuery);
                                querySetCell(local.excelOutputQuery, "FirstName", local.row.FIRST_NAME);
                                querySetCell(local.excelOutputQuery, "LastName", local.row.LAST_NAME);
                                querySetCell(local.excelOutputQuery, "Address", local.row.ADDRESS);
                                querySetCell(local.excelOutputQuery, "Email", local.row.EMAIL);
                                querySetCell(local.excelOutputQuery, "Phone", local.row.PHONE);
                                querySetCell(local.excelOutputQuery, "DOB", local.row.DOB);
                                querySetCell(local.excelOutputQuery, "Role", local.row.ROLE);
                            }
                        
                        objSpreadsheet          =   SpreadsheetNew("Sheet1",true);
                        SpreadsheetAddRow(objSpreadsheet, "First Name,Last Name,Address,Email,Phone,DOB,Role" );
                        SpreadsheetFormatRow(objSpreadsheet, {bold=true, alignment="center"}, 1 );
                        spreadsheetAddRows(objSpreadsheet, local.excelOutputQuery);
                        cfheader(
                                    name="Content-Disposition",
                                    value="attachment; filename=Template_with_data.xlsx"
                                );
                        cfcontent(
                                    type="application/vnd.ms-excel.sheet.macroEnabled.12",
                                    variable="#SpreadsheetReadBinary( objSpreadsheet )#"
                                );
                    }
                else 
                    {
                        writeOutput('Invalid Type');
                    }
            }

        public function processMyExcel(required query excelQuery)
            {
                local.hasData                   =   false;
                local.excelOutputQuery          =   queryNew("FirstName, LastName, Address, Email, Phone, DOB, Role, Result, Flag");

                for(row IN arguments.excelQuery)
                    {
                        local.checkEmptyExcel = checkIfEmptyExcel(row);
                        if(local.checkEmptyExcel)
                            {
                                local.hasData           = true;
                                local.validateColumns   = checkEmptyColumns(row);
                                local.forQueryStruct    = structNew();
                                
                                if(local.validateColumns === 'success')
                                    {
                                        local.addNewUser                    =   processUser(row['First Name'], row['Last Name'], row['Address'], row['Email'], row['Phone'], row['DOB'], row['Role']);
                                        queryAddRow(local.excelOutputQuery);
                                        querySetCell(local.excelOutputQuery, "FirstName", row['First Name']);
                                        querySetCell(local.excelOutputQuery, "LastName", row['Last Name']);
                                        querySetCell(local.excelOutputQuery, "Address", row['Address']);
                                        querySetCell(local.excelOutputQuery, "Email", row['Email']);
                                        querySetCell(local.excelOutputQuery, "Phone", row['Phone']);
                                        querySetCell(local.excelOutputQuery, "DOB", row['DOB']);
                                        querySetCell(local.excelOutputQuery, "Role", row['Role']);
                                        
                                        if(local.addNewUser == 0)
                                            {
                                                querySetCell(local.excelOutputQuery, "Result", 'Added');
                                                querySetCell(local.excelOutputQuery, "Flag", '0');
                                            }
                                        else if(local.addNewUser == 1)
                                            {
                                                querySetCell(local.excelOutputQuery, "Result", 'Updated');
                                                querySetCell(local.excelOutputQuery, "Flag", '1');
                                            }
                                        else 
                                            {
                                                querySetCell(local.excelOutputQuery, "Result", 'Database updation failed');
                                                querySetCell(local.excelOutputQuery, "Flag", '2');
                                            } 
                                    }
                                else 
                                    {
                                        queryAddRow(local.excelOutputQuery);
                                        querySetCell(local.excelOutputQuery, "FirstName", row['First Name']);
                                        querySetCell(local.excelOutputQuery, "LastName", row['Last Name']);
                                        querySetCell(local.excelOutputQuery, "Address", row['Address']);
                                        querySetCell(local.excelOutputQuery, "Email", row['Email']);
                                        querySetCell(local.excelOutputQuery, "Phone", row['Phone']);
                                        querySetCell(local.excelOutputQuery, "DOB", row['DOB']);
                                        querySetCell(local.excelOutputQuery, "Role", row['Role']);
                                        querySetCell(local.excelOutputQuery, "Result", local.validateColumns);
                                        querySetCell(local.excelOutputQuery, "Flag", '2');
                                    }
                            }      
                    }

                if(local.hasData == false)
                    {
                        return 'empty_excel';
                    }
                else 
                    {
                        local.sortedQuery =   QuerySort(local.excelOutputQuery,function(obj1,obj2){
                            return compare(obj2.Flag,obj1.Flag)
                        });
                        QueryDeleteColumn(local.excelOutputQuery,"Flag");
                        objSpreadsheet          = SpreadsheetNew("Sheet1",true);
                        SpreadsheetAddRow( objSpreadsheet, "First Name, Last Name, Address, Email, Phone, DOB, Role, Result" );
                        SpreadsheetFormatRow( objSpreadsheet, {bold=true, alignment="center"}, 1 );
                        spreadsheetAddRows(objSpreadsheet, local.excelOutputQuery);
                        local.filename = expandPath("./uploads/Upload_Result.xlsx")    
                        spreadsheetWrite(objSpreadsheet, local.filename, true);
                        return 'success';
                    }
            }

        public function getUsers()
            {
                try 
                    {
                        local.result = queryExecute("SELECT * FROM users");
                        return local.result;	
                    }
                catch(Exception e) 
                    {
                        return 'error';
                    }
            }

        public function getRoles()
            {
                try
                    {
                        local.result = queryExecute("SELECT GROUP_CONCAT(roles) AS roles FROM roles",{returntype="array"});
                        return local.result;
                    }
                catch(Exception e)
                    {
                        return 'error';
                    }
            }

        public function getEmail(required string email)
            {
                try 
                    {
                        local.result = queryExecute("SELECT COUNT(*) AS total FROM users WHERE email = :email",
                                                        {email: {cfsqltype: "cf_sql_varchar", value: arguments.email}},
                                                        {returntype="array"}
                                                    );
                        return local.result;	
                    }
                catch(Exception e) 
                    {
                        return 'error';
                    }
            }

        private function checkEmptyColumns(any row)
            {
                local.hasError    = false;
                local.resultArray = arrayNew(1, true);
                if(len(trim(arguments.row['First Name'])) === 0)
                    {
                        local.hasError          =   true;
                        local.resultArray.append('First Name is missing');
                    }

                if(len(trim(arguments.row['Last Name'])) === 0)
                    {
                        local.hasError          =   true;
                        local.resultArray.append('Last Name is missing');
                    }

                if(len(trim(arguments.row['Email'])) === 0)
                    {
                        local.hasError          =   true;
                        local.resultArray.append('Email is missing');
                    }

                if(len(trim(arguments.row['Phone'])) === 0)
                    {
                        local.hasError          =   true;
                        local.resultArray.append('Phone is missing');
                    }    
                
                if(len(trim(arguments.row['DOB'])) === 0)
                    {
                        local.hasError          =   true;
                        local.resultArray.append('DOB is missing');
                    }

                if(len(trim(arguments.row['Address'])) === 0)
                    {
                        local.hasError          =   true;
                        local.resultArray.append('Address is missing');
                    }    

                if(len(trim(arguments.row['Role'])) === 0)
                    {
                        local.hasError          =   true;
                        local.resultArray.append('Role is missing');
                    }
                else 
                    {
                        local.roleGroupSet      =   getRoles();
                        if(NOT findNoCase(arguments.row['Role'], local.roleGroupSet.roles))
                            {
                                local.hasError          =   true;
                                local.resultArray.append('Role is not valid');
                            }
                    }

                if(local.hasError === false)
                    {
                        local.resultArray.append('success');
                    }

                return arrayToList(local.resultArray);
            }

        private function checkIfEmptyExcel(any row)
            {
                if(len(trim(arguments.row['First Name'])) > 0 
                    OR len(trim(arguments.row['Last Name'])) > 0  
                        OR len(trim(arguments.row['Email'])) > 0 
                            OR len(trim(arguments.row['Phone'])) > 0
                                OR len(trim(arguments.row['DOB'])) > 0
                                    OR len(trim(arguments.row['Address'])) > 0
                                        OR len(trim(arguments.row['Role'])) > 0)
                                            {
                                                return true;
                                            }
                                        else 
                                            {
                                                return false;
                                            }
            }

        private function processUser(first_name, last_name, address, email, phone, dob, role)
            {
                try
                    {
                        local.duplicateEmail      =   getEmail(arguments.email);
                        if(local.duplicateEmail[1].total > 0)
                            {
                                result = queryExecute("UPDATE users 
                                                                SET 
                                                                    first_name  = :first_name, 
                                                                    last_name   = :last_name, 
                                                                    address     = :address, 
                                                                    phone       = :phone, 
                                                                    dob         = :dob
                                                                WHERE   
                                                                    email = :email",
                                                            {
                                                                email: { cfsqltype: "cf_sql_varchar", value: arguments.email },
                                                                first_name: { cfsqltype: "cf_sql_varchar", value: arguments.first_name },
                                                                last_name: { cfsqltype: "cf_sql_varchar", value: arguments.last_name },
                                                                address: { cfsqltype: "cf_sql_varchar", value: arguments.address },
                                                                phone: { cfsqltype: "cf_sql_varchar", value: arguments.phone },
                                                                dob: { cfsqltype: "cf_sql_date", value: arguments.dob }
                                                            }, 
                                                            { result="resultset" });
                                return	1;
                            }
                        else 
                            {
                                result = queryExecute("INSERT INTO users (
                                                                    first_name, 
                                                                    last_name, 
                                                                    address, 
                                                                    email, 
                                                                    phone, 
                                                                    dob
                                                                ) 
                                                        VALUES 
                                                            (	
                                                                :first_name,
                                                                :last_name,
                                                                :address,
                                                                :email,
                                                                :phone,
                                                                :dob
                                                            )",
                                                            {
                                                                first_name: { cfsqltype: "cf_sql_varchar", value: arguments.first_name },
                                                                last_name: { cfsqltype: "cf_sql_varchar", value: arguments.last_name },
                                                                address: { cfsqltype: "cf_sql_varchar", value: arguments.address },
                                                                email: { cfsqltype: "cf_sql_varchar", value: arguments.email },
                                                                phone: { cfsqltype: "cf_sql_varchar", value: arguments.phone },
                                                                dob: { cfsqltype: "cf_sql_date", value: arguments.dob }
                                                            }, 
                                                            { result="resultset" });
                                return	0;
                            }                        
                    }
                catch(Exception e)
                    {
                        return 'error';
                    }
            }

    }