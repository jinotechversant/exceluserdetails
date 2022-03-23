component displayname="userdetails"
    {
        public function processMyExcel(excelQuery)
            {
                local.hasData   =   false;

                for(row IN excelQuery)
                    {
                        local.checkEmptyExcel = checkIfEmptyExcel(row);
                        if(local.checkEmptyExcel)
                            {
                                local.hasData           = true;
                                local.validateColumns   = checkEmptyColumns(row);
                                if(local.validateColumns === 'success')
                                    {
                                        local.addNewUser   = addUser(row['First Name'], row['Last Name'], row['Address'], row['Email'], row['Phone'], row['DOB'], row['Role']);
                                    }
                                else 
                                    {

                                    }
                            }      
                    }


                if(local.hasData == false)
                    {
                        return 'empty_excel';
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

        public function getEmail(email)
            {
                try 
                    {
                        local.result = queryExecute("SELECT COUNT(*) AS total FROM users WHERE email = :email",
                                                        {email: {cfsqltype: "cf_sql_varchar", value: email}},
                                                        {returntype="array"}
                                                    );
                        return local.result;	
                    }
                catch(Exception e) 
                    {
                        return 'error';
                    }
            }

        private function checkEmptyColumns(row)
            {
                local.hasError    = false;
                local.resultArray = arrayNew(1, true);
                if(len(trim(row['First Name'])) === 0)
                    {
                        local.hasError          =   true;
                        local.resultArray.append('First Name is missing');
                    }

                if(len(trim(row['Last Name'])) === 0)
                    {
                        local.hasError          =   true;
                        local.resultArray.append('Last Name is missing');
                    }

                if(len(trim(row['Email'])) === 0)
                    {
                        local.hasError          =   true;
                        local.resultArray.append('Email is missing');
                    }
                else 
                    {
                        local.duplicateEmail      =   getEmail(row['Email']);
                        if(local.duplicateEmail[1].total > 0)
                            {
                                local.hasError          =   true;
                                local.resultArray.append('Duplicate email entry found');
                            }
                    }

                if(len(trim(row['Phone'])) === 0)
                    {
                        local.hasError          =   true;
                        local.resultArray.append('Phone is missing');
                    }    
                
                if(len(trim(row['DOB'])) === 0)
                    {
                        local.hasError          =   true;
                        local.resultArray.append('DOB is missing');
                    }

                if(len(trim(row['Address'])) === 0)
                    {
                        local.hasError          =   true;
                        local.resultArray.append('Address is missing');
                    }    

                if(len(trim(row['Role'])) === 0)
                    {
                        local.hasError          =   true;
                        local.resultArray.append('Role is missing');
                    }
                else 
                    {
                        local.roleGroupSet      =   getRoles();
                        if(NOT findNoCase(row['Role'], local.roleGroupSet.roles))
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

        private function checkIfEmptyExcel(row)
            {
                if(len(trim(row['First Name'])) > 0 
                    OR len(trim(row['Last Name'])) > 0  
                        OR len(trim(row['Email'])) > 0 
                            OR len(trim(row['Phone'])) > 0
                                OR len(trim(row['DOB'])) > 0
                                    OR len(trim(row['Address'])) > 0
                                        OR len(trim(row['Role'])) > 0)
                                            {
                                                return true;
                                            }
                                        else 
                                            {
                                                return false;
                                            }
            }

        private function addUser(first_name, last_name, address, email, phone, dob, role)
            {
                try
                    {
                        result = queryExecute("INSERT INTO users (
                                                            first_name, 
                                                            last_name, 
                                                            address, 
                                                            email, 
                                                            phone, 
                                                            dob,
                                                            role
                                                        ) 
                                                VALUES 
                                                    (	
                                                        :first_name,
                                                        :last_name,
                                                        :address,
                                                        :email,
                                                        :phone,
                                                        :dob,
                                                        :role
                                                    )",
                                                    {
                                                        first_name: { cfsqltype: "cf_sql_varchar", value: first_name },
                                                        last_name: { cfsqltype: "cf_sql_varchar", value: last_name },
                                                        address: { cfsqltype: "cf_sql_varchar", value: address },
                                                        email: { cfsqltype: "cf_sql_varchar", value: email },
                                                        phone: { cfsqltype: "cf_sql_varchar", value: phone },
                                                        dob: { cfsqltype: "cf_sql_date", value: dob },
                                                        role: { cfsqltype: "cf_sql_varchar", value: role }
                                                    }, 
                                                    { result="resultset" });
                        return	resultset.generatedKey;
                    }
                catch(Exception e)
                    {
                        return 'error';
                    }
            }

    }