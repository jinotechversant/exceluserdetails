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
                                writeDump(local.validateColumns);
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
                        local.result = queryExecute("SELECT GROUP_CONCAT(roles) AS roles FROM roles");
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
                local.resultArray = arrayNew(1);
                if(len(trim(row['First Name'])) === 0)
                    {
                        local.hasError          =   true;
                        local.resultArray[1]    =   'First Name is missing';
                    }

                if(len(trim(row['Last Name'])) === 0)
                    {
                        local.hasError          =   true;
                        local.resultArray[1]    =   'Last Name is missing';
                    }

                if(len(trim(row['Email'])) === 0)
                    {
                        local.hasError          =   true;
                        local.resultArray[1]    =   'Email is missing';
                    }

                if(len(trim(row['Phone'])) === 0)
                    {
                        local.hasError          =   true;
                        local.resultArray[1]    =   'Phone is missing';
                    }    
                
                if(len(trim(row['DOB'])) === 0)
                    {
                        local.hasError          =   true;
                        local.resultArray[1]    =   'DOB is missing';
                    }

                if(len(trim(row['Address'])) === 0)
                    {
                        local.hasError          =   true;
                        local.resultArray[1]    =   'Address is missing';
                    }    

                if(len(trim(row['Role'])) === 0)
                    {
                        local.hasError          =   true;
                        local.resultArray[1]    =   'Role is missing';
                    }

                if(local.hasError === false)
                    {
                        local.resultArray[1]    =   'success';
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

    }