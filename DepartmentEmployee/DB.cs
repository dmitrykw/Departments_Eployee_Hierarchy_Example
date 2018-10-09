using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;


namespace DepartmentEmployee
{
    class DB
    {


        private MSSQLConnectionParams sqlconn_params;
        private SqlConnection conn = new SqlConnection(); //Объявляем переменную нашего Connection

        public DB(MSSQLConnectionParams sqlconn_params) //В конструкторе при создании экземпляра класса получаем на вход параметры SQL соединеия нашего типа SQLConnectionParams
        {
            this.sqlconn_params = sqlconn_params;
        }



        public async Task<bool> ExecSQLAsync(string sqlCommand) //Метод для выполнения произвольной SQL команды
        {
            return await Task.Run(() =>
            {
                    if (conn.State != ConnectionState.Open) //Открываем соединение
                {
                    ConnectionOpen();
                }

                try
                {
                    SqlCommand cmd = new SqlCommand(sqlCommand, conn);
                    cmd.CommandType = CommandType.Text;
                    cmd.ExecuteNonQuery();

                    if (conn.State != ConnectionState.Closed)
                    {
                        ConnectionClose(); //Закрываем соединение                   
                    }

                    return true;
                }
                catch (Exception ex){ MessageBox.Show("Error executing MSsql Command: " + ex.Message);

                    if (conn.State != ConnectionState.Closed)
                    {
                        ConnectionClose(); //Закрываем соединение                   
                    }

                    return false;
                }

                
            });
        }

        //Метод для выполнения SQL команды - перегруженный метод для параметризированных запросов. Принимаем коллекцию ключ значение содержащую название параметра (например @name)и значение параметра (например "John Smith")
        public async Task<bool>  ExecSQLAsync(string sqlCommand, Dictionary<string, string> mssql_parameters_list)
        {

            return await Task.Run(() =>
            {

                if (conn.State != ConnectionState.Open)
            {
                ConnectionOpen();
            }

            try
            {
                SqlCommand cmd = new SqlCommand(sqlCommand, conn);
                cmd.CommandType = CommandType.Text;

                //Перебираем коллекцию параметров и добавляем каждый 
                foreach (KeyValuePair<string, string> mssql_parameter in mssql_parameters_list)
                {
                    //аргументы - название параметра, тип данных в поле SQL, длинна строки. Value - соответсвенно значение.
                    cmd.Parameters.Add(new SqlParameter(mssql_parameter.Key, SqlDbType.VarChar, 255) { Value = mssql_parameter.Value });
                }

                cmd.ExecuteNonQuery();

                    if (conn.State != ConnectionState.Closed)
                    {
                        ConnectionClose(); //Закрываем соединение                   
                    }

                    return true;

                }
                catch (Exception ex){MessageBox.Show("Error executing MSsql Command: " + ex.Message);

                    if (conn.State != ConnectionState.Closed)
                    {
                        ConnectionClose(); //Закрываем соединение                   
                    }

                    return false;
                }

            
            });

        }








        private void ConnectionOpen() //Метод для открытия соединения
        {

            try
            {
                string conString = @"Data Source=" + sqlconn_params.hostname + ";Initial Catalog=" + sqlconn_params.database + ";integrated security=False;User ID=" + sqlconn_params.user + ";Password=" + sqlconn_params.passwd;
                conn = new SqlConnection(conString);


                if (conn.State != ConnectionState.Open)
                {
                    conn.Open();// Открываем соединиение
                }

            }
            catch (Exception ex ){ MessageBox.Show("Error to establish the MSsql Connection: " + ex.Message); }
        }

        private void ConnectionClose() //Метод для зарытия соединения
        {
            try
            {
                if (conn.State != ConnectionState.Closed)
                {
                    conn.Close(); //Закрываем соединение                   
                }
            }
            catch (Exception ex) { MessageBox.Show("Error to close MSSql Connection: "+ ex.Message); }
        }




        //Метод извлечения daratable из базы данных MSSQL Async
        public async Task<DataTable> GetDatatableFromMSSQLAsync(string sqlcommand)
        {

            return await Task.Run(() =>
            {
                if (conn.State != ConnectionState.Open)
                {
                    ConnectionOpen();
                }

                 try
                 {

                SqlCommand cmd = new SqlCommand(sqlcommand, conn);

                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);

                    DataTable results = new DataTable();


                    sda.Fill(results);

              
                //Удаляем первый столбец - там ID шник
                //results.Columns.RemoveAt(0);

                if (conn.State != ConnectionState.Closed)
                {
                    ConnectionClose(); //Закрываем соединение                   
                }

                return results;
                }
                catch
                 {
                   DataTable results = new DataTable();
                   return results;
                 }


            });
            
        }

        //Метод извлечения daratable из базы данных MSSQL
        public DataTable GetDatatableFromMSSQL(string sqlcommand)
        {
            
                if (conn.State != ConnectionState.Open)
                {
                    ConnectionOpen();
                }

                try
                {

                    SqlCommand cmd = new SqlCommand(sqlcommand, conn);

                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);

                    DataTable results = new DataTable();


                    sda.Fill(results);


                    //Удаляем первый столбец - там ID шник
                    //results.Columns.RemoveAt(0);

                    if (conn.State != ConnectionState.Closed)
                    {
                        ConnectionClose(); //Закрываем соединение                   
                    }

                    return results;
                }
                catch
                {
                    DataTable results = new DataTable();
                    return results;
                }



        }

    }
}
