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
                                local.hasData = true;
                                writeDump(row)
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
                
            }

        private function checkIfEmptyExcel(row)
            {
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