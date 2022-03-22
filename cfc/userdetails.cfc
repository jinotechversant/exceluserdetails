component displayname="userdetails"
    {
        public function processExcel(excelQuery)
            {
                for(row IN excelQuery)
                    {
                        local.checkEmptyValues = checkEmpty(row);
                        if(local.checkEmptyValues)
                            {
                               writeDump(row['Address'])
                            }
                        else 
                            {
                                writeOutput('error');
                            }       
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

        private function checkEmpty(row)
            {
                /*Checks for empty rows*/

                if(len(trim(row['First Name'])) > 0 
                    AND len(trim(row['Last Name'])) > 0  
                        AND len(trim(row['Email'])) > 0 
                            AND len(trim(row['DOB'])) > 0
                                AND len(trim(row['Address'])) > 0
                                    AND len(trim(row['Role'])) > 0)
                                        {
                                            return true;
                                        }
                                    else 
                                        {
                                            return false;
                                        }
            }

    }