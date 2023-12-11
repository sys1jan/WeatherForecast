/* 
public string TranslateBayToPosition(string bayNr)
        {
            string bayToPosition = "";
            try
            {
                string sql;
 
                sql = "SELECT POS_NR FROM TLINUPT WHERE BAY_NR = '" + bayNr + "'";
 
                using (IDataReader bayReader = _dbConn.ExecuteReader(sql))
                {
                    if (bayReader != null)
                    {
                        if (bayReader.Read())
                        {
                            bayToPosition = bayReader["POS_NR"].ToString();
                        }
                        else
                        {
                            bayToPosition = bayNr;
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                Logging.Error(ex);
                return string.Empty;
            }
 
            return bayToPosition;
        }
 */