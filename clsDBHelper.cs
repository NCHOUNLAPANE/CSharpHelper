using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.VisualBasic;
using System.Data.SqlClient;
using System.Data;
using System.Windows.Forms;

namespace MachineStatus
{

    public class clsDBHelper
    {
        
        public clsDBHelper(SqlConnection conn)
        {
            _conn = conn;
        }
        
        private SqlConnection _conn;

        public System.Data.DataTable DataTbl = new System.Data.DataTable();


        /// <summary>
        /// Executes a stored proc and returns a data table with the results.
        /// </summary>
        /// <param name="strStoredProc"></param>
        /// <param name="_parameters"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public DataTable ExecuteStoredProcDT(string strStoredProc, SqlParameter[] _parameters)
        {


            try
            {
                
                SqlCommand cm = new SqlCommand(strStoredProc, _conn);
                cm.CommandType = CommandType.StoredProcedure;

                // Check to see if there are any paramters being passed in.
                // If there are none, then do not add and execute the SP
                if ((((_parameters) != null)))
                {
                    foreach (SqlParameter p in _parameters)
                    {
                        cm.Parameters.Add(p);
                    }
                }

                SqlDataAdapter da = new SqlDataAdapter(cm);
                DataSet ds = new DataSet();
                DataTable dt = new DataTable();

                da.Fill(ds, "rs");
                dt = ds.Tables["rs"];
                return dt;



            }
            catch (Exception ex)
            {
                throw new ApplicationException(ex.Message);

              //  return null;

            }
        }


        /// <summary>
        /// Executes a stored proc and returns a data set - usefull for when stored procs return multiple recordsets.
        /// </summary>
        /// <param name="strStoredProc"></param>
        /// <param name="_parameters"></param>
        /// <returns></returns>
        public DataSet ExecuteStoredProcDS(string strStoredProc, SqlParameter[] _parameters)
        {

            try
            {
                SqlCommand cm = new SqlCommand(strStoredProc, _conn);
                cm.CommandType = CommandType.StoredProcedure;
                // Check to see if there are any paramters being passed in.
                // If there are none, then do not add and execute the SP
                if ((((_parameters) != null)))
                {
                    foreach (SqlParameter p in _parameters)
                    {
                        cm.Parameters.Add(p);
                    }
                }
                SqlDataAdapter da = new SqlDataAdapter(cm);
                DataSet ds = new DataSet();

                da.Fill(ds);
                return ds;


            }
            catch (Exception ex)
            {
                throw new ApplicationException(ex.Message);
                // return null;
            }
        }

        public bool ExecuteStoredProc(string strStoredProc, SqlParameter[] _parameters)
        {
            try
            {
               
                    SqlCommand cm = new SqlCommand(strStoredProc, _conn);
                cm.CommandType = CommandType.StoredProcedure;

                if ((((_parameters) != null)))
                {
                    foreach (SqlParameter p in _parameters)
                    {
                        cm.Parameters.Add(p);

                    }
                }

                cm.ExecuteNonQuery();

                return true;


            }
            catch (Exception ex)
            {
                throw new ApplicationException(ex.Message);


            }
        }


        public bool GetRecordset(string strQuery)
        {
            SqlDataAdapter da;
            if (!(DataTbl == null))
            {
                DataTbl.Clear();
                DataTbl.Rows.Clear();
                DataTbl.Columns.Clear();
            }
            try
            {
                da = new SqlDataAdapter(strQuery, _conn);
                da.Fill(DataTbl);
                return true;
            }
            catch (SqlException ex)
            {
                MessageBox.Show("Error Getting Recordset: " + ex.Message); 
                return false;
            }
        }

        public System.Data.DataTable GetRecordsetDT(string strQuery)
        {
            SqlDataAdapter da;

            try
            {
                System.Data.DataTable dt = new System.Data.DataTable();
                da = new SqlDataAdapter(strQuery, _conn);
                da.Fill(dt);
                return dt;
            }
            catch (SqlException ex)
            {
                MessageBox.Show("Error Getting RecordsetDT: " + ex.Message); 
                return null;
            }
        }

        public DataTable GetDataTable(string strQuery)
        {
            try { 
                DataTable dataTable = new DataTable();

                SqlCommand cmd = _conn.CreateCommand();
                cmd.CommandText = strQuery;
                cmd.CommandType = CommandType.Text;

                // SqlDataReader need an open conncetion, so check and open it.
                if (_conn.State != ConnectionState.Open)
                    _conn.Open();

                // Read data by using Execute Reader
                SqlDataReader dr = cmd.ExecuteReader();//(CommandBehavior.CloseConnection);

                //Use data table load method to load data from data reader
                dataTable.Load(dr);

                return dataTable;
            }
            catch (SqlException ex)
            {
                MessageBox.Show("Error Getting RecordsetDT: " + ex.Message);
                return null;
            }
        }



        public void FillGrid(ref System.Windows.Forms.DataGridView grid, ref DataTable DataTable, params object[] strFieldList)
        {
            try
            {
                int x;
                int i;
                object[] objCells = new object[strFieldList.Length];
                grid.Rows.Clear();

                for (x = 0; x <= DataTable.Rows.Count - 1; x++)
                {
                    for (i = 0; (i <= strFieldList.Length-1); i++)
                    {
                        if ((strFieldList[i].ToString().Trim() != ""))
                        {
                            objCells[i] = Nvl(DataTable.Rows[x][strFieldList[i].ToString()],"");
                        }
                    }
                    grid.Rows.Add(objCells);
                }
            }
            catch (Exception ex)
            {
             //Currently not using FillGrid for anything yet.
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        public object Nvl(object vField, object vReplaceWith)
        {


            if (vField == System.DBNull.Value | (vField == null))
                return vReplaceWith;

            else
                return vField;



        }
        public bool ExecuteNonQuery(string strQuery)
        {
            SqlCommand objCommand;

            try
            {

                objCommand = new SqlCommand(strQuery, _conn);

               

                objCommand.ExecuteNonQuery();
                return true;
            }
            catch (SqlException ex)
            {
                MessageBox.Show("Error Executing Non-Query to SQL: " + ex.Message); 
                return false;
            }
            finally
            {
                objCommand = null;
            }
        }

        public bool ExecuteNonQuery(string strQuery, int iCommandTimeout)
        {
            SqlCommand objCommand;

            try
            {

                objCommand = new SqlCommand(strQuery, _conn);

                //Use an increased timeout
                objCommand.CommandTimeout = iCommandTimeout;

                objCommand.ExecuteNonQuery();
                return true;
            }
            catch (SqlException ex)
            {
                MessageBox.Show("Error Executing NonQuery to SQL: " + ex.Message); 
                return false;
            }
            finally
            {
                objCommand = null;
            }
        }


        public string BuildUpdateQry(string strTableName, params object[] objFieldsAndValues)
        {
            //Function returns a formatted insert query only. Makes NO connection to the database. 

            string qry;
            Int16 x;
            qry = "Update " + strTableName + " Set ";
            for (x = 0; x <= objFieldsAndValues.GetLength(0) - 1; x += 2)
            {

                if (objFieldsAndValues[x + 1] == null)
                    objFieldsAndValues[x + 1] = "";

                string str = objFieldsAndValues[x + 1].GetType().Name;
                str = str.ToUpper();
                switch (str)
                {
                    case "BOOLEAN":
                        if ((bool)objFieldsAndValues[x + 1] == true)
                            qry += objFieldsAndValues[x] + "= 1, ";
                        else
                            qry += objFieldsAndValues[x] + "= 0, ";
                        break;
                    case "INT32":
                    case "INT16":
                    case "INT64":
                        qry += objFieldsAndValues[x] + "=" + objFieldsAndValues[x + 1] + ", ";
                        break;
                    case "STRING":
                    case "DATETIME":
                        if (objFieldsAndValues[x + 1].ToString() == "NULL")
                            qry += objFieldsAndValues[x] + "= NULL, ";
                        else
                            qry += objFieldsAndValues[x] + "= '" + FixForApostrophe(objFieldsAndValues[x + 1].ToString()) + "', ";

                        break;
                    default:
                        qry += objFieldsAndValues[x] + "=" + objFieldsAndValues[x + 1] + ", ";
                        break;
                }

            }

            return qry.Substring(0, qry.Length - 2);
        }

        /// <summary>
        /// Note - FixForApostrophe is already built into this query generator
        /// </summary>
        /// <param name="strTableName"></param>
        /// <param name="objFieldsAndValues"></param>
        /// <returns></returns>
        public string BuildInsertQry(string strTableName, params object[] objFieldsAndValues)
        {
            // Function returns a formatted insert query only.  Makes NO connection to the database.

            Int16 x;
            string qry;
            qry = ("Insert Into " + (strTableName + " ("));
            for (x = 0; x <= objFieldsAndValues.GetLength(0) - 1; x += 2)
            {
                qry = (qry + (objFieldsAndValues[x] + ","));
            }
            qry = (qry.Substring(0, (qry.Length - 1)) + ") VALUES (");
            for (x = 1; x < objFieldsAndValues.GetLength(0); x += 2)
            {
                if ((objFieldsAndValues[x] == null))
                {
                    qry = (qry + (objFieldsAndValues[x] + "\'\',"));
                }
                else
                {
                    string str = objFieldsAndValues[x].GetType().Name;
                    str = str.ToUpper();
                    switch (str)
                    {
                        case "BOOLEAN":
                            if ((bool)objFieldsAndValues[x] == true)
                                qry = qry + "1,";
                            else
                                qry = qry + "0,";

                            break;
                        case "INT64":
                        case "INT32":
                        case "INT16":
                            qry = (qry + (objFieldsAndValues[x] + ","));
                            break;
                        case "STRING":
                            qry = (qry + ("\'" + (FixForApostrophe((string)objFieldsAndValues[x]) + "\',")));
                            break;
                        case "DATETIME":
                            qry = (qry + ("\'" + (objFieldsAndValues[x] + "\',")));
                            break;
                        default:
                            qry = (qry + (objFieldsAndValues[x] + ","));
                            break;
                    }
                }
            }
            return qry.Substring(0, (qry.Length - 1)) + ")";


        }


        public string FixForApostrophe(string strFieldContents)
        {
            //  replace each apostrophe with two apostrophes
            return strFieldContents.Replace("'", "''");

        }

        
        /// <summary>
        ///  Use ExecuteScalar method to return a single value from the database. 
        ///  The function returns the field object value
        /// </summary>
        /// <param name="strQuery"></param>
        /// <returns></returns>
        public object ExecuteScalar(string strQuery)
        {
            
            SqlCommand objCommand;
            try
            {
                

              
                objCommand = _conn.CreateCommand();
                objCommand.CommandText = strQuery;
                return objCommand.ExecuteScalar();
            }
            catch (SqlException  ex)
            {
             
                System.Windows.Forms.MessageBox.Show(ex.Message);
                return null;
            }
            finally
            {
                objCommand = null;
            }
        }

        /*

        /// <summary>
        /// Returns a SQLDataReader filled with one or more query results.  Separate multiple queries with semi-colons. 
        /// Use DataReader.NextResult to move between resultsets.
        /// You MUST close the reader when you're done with it.
        /// </summary>

        public SqlDataReader DataReader(string strQueries)
        {
            SqlCommand cmd;
            SqlConnection conn = new SqlConnection();
            conn.ConnectionString = _strConnectionString;
            conn.Open();
            cmd = new SqlCommand(strQueries, conn);
            return cmd.ExecuteReader();
        }

     
        

  


   

     

       
        */


    }

}