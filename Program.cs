using System;
using System.IO;
using System.Net;
using System.Data.OleDb;
using System.Data;
using System.Data.SqlClient;

namespace NBKURSY
{
    class Program
    {
        
        static void Main(string[] args)
        {
            ///-------------------------USD----------------------------------------------------------------
            DateTime datin = DateTime.Now.AddDays(1);
            string YYYin = datin.Year.ToString();
            string MMMin = datin.Month.ToString();
            string DDDin = datin.Day.ToString();
            string datYMDin = YYYin + "-" + MMMin + "-" + DDDin;

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create("https://www.nbrb.by/api/exrates/rates/USD?parammode=2&ondate=" + datYMDin);
            string poisk = "Cur_OfficialRate";
            string poiskDate = "Date";
            string writePath = "log.txt";
            try
            {
                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                {
                    using (Stream stream = response.GetResponseStream())
                    {
                        using (StreamReader reader = new StreamReader(stream))
                        {
                            string str = reader.ReadToEnd();
                            Console.WriteLine(str);
                            
                            int number1 = str.IndexOf(poisk);
                            string kursUSD = str.Substring(number1 + 18, 6);


                            int numberDate = str.IndexOf(poiskDate);
                            string dateUSD = str.Substring(numberDate + 7, 10);
                            DateTime dat = Convert.ToDateTime(dateUSD);
                            string YYY = dat.Year.ToString();
                            string MMM = dat.Month.ToString();
                            string DDD = dat.Day.ToString();
                            string datYMD = YYY + "-" +DDD + "-" + MMM;
                           
                            

                            ///----------------------------------------------------------------
                            string connectionString = "Provider = SQLOLEDB.1; Persist Security Info = True; User ID = sa; Initial Catalog = General; Data Source = galsrv\\ins1; Password = Cd3pk7zr% ";
                            using (OleDbConnection connection = new OleDbConnection(connectionString))
                            {                               
                                string comm1 = "select dbo.ToDate(f$DATVAL), F$SUMRUBL from t$cursval where f$DATVAL = dbo.ToAtlDate('" + datYMD + "') and f$KODVALUT = 0x8000000000000002";
                                string comm2 = "insert into t$cursval(f$KODVALUT, f$DATVAL, F$SUMRUBL, F$CMAIN) values(0x8000000000000002, dbo.ToAtlDate('" + datYMD + "'), " + kursUSD + ", 0x8000000000000002)";

                                OleDbCommand command1 = new OleDbCommand(comm1);
                                OleDbCommand command2 = new OleDbCommand(comm2);

                                command1.Connection = connection;
                                command2.Connection = connection;
                                
                                // Open the connection and execute the insert command.
                                try
                                {   
                                    connection.Open();
                                    command1.ExecuteNonQuery();
                                    
                                    OleDbDataReader reader1 = command1.ExecuteReader();
                                    if (reader1.HasRows) // если есть данные
                                    {
                                        using (StreamWriter sw = new StreamWriter(writePath, true, System.Text.Encoding.Default))
                                        {
                                            sw.Write(DateTime.Now);
                                            sw.WriteLine(" Обнаружен в базе курс Доллара США: ");
                                            
                                            sw.WriteLine("Дата\t\tКурс");

                                            while (reader1.Read()) // построчно считываем данные
                                        {
                                            object dats = reader1.GetValue(0);
                                            object kurs = reader1.GetValue(1);
                                            
                                            sw.WriteLine("{0} \t{1}", dats, kurs);
                                            
                                        }
                                            sw.WriteLine("--------------------------------------------------------------------------------------------------------------");
                                            
                                        }                                                                                
                                    }

                                    else 
                                    {
                                        command2.ExecuteNonQuery();
                                        connection.Close();
                                        using (StreamWriter sw = new StreamWriter(writePath, true, System.Text.Encoding.Default))
                                        {
                                            sw.Write(DateTime.Now);
                                            sw.Write(" Введен курс Доллара США: ");
                                            sw.Write(kursUSD + " на дату: ");
                                            sw.WriteLine(dateUSD);
                                        }
                                    }
                                    
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine(ex.Message);
                                    using (StreamWriter sw = new StreamWriter(writePath, true, System.Text.Encoding.Default))
                                    {
                                        sw.WriteLine("Ошибка");
                                        sw.Write(DateTime.Now + ": ");
                                        sw.WriteLine(ex.Message);
                                        Console.WriteLine("Значение dat======:" + dat);
                                        Console.WriteLine("Значение datYMD======:" + datYMD);
                                       
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                using (StreamWriter sw = new StreamWriter(writePath, true, System.Text.Encoding.Default))
                {
                    sw.WriteLine("Ошибка");
                    sw.Write(DateTime.Now + ": ");
                    sw.WriteLine(ex.Message);
                }
            }
            //-------------------EUR-----------------------------------------------------------------------------
            try
            {
                HttpWebRequest requestEUR = (HttpWebRequest)WebRequest.Create("https://www.nbrb.by/api/exrates/rates/EUR?parammode=2&ondate=" + datYMDin);

                //string writePath = @"D:\kurs.txt";
                string poiskEUR = "Cur_OfficialRate";
                string poiskDateEUR = "Date";

                using (HttpWebResponse responseEUR = (HttpWebResponse)requestEUR.GetResponse())
                {
                    using (Stream streamEUR = responseEUR.GetResponseStream())
                    {
                        using (StreamReader readerEUR = new StreamReader(streamEUR))
                        {
                            string str = readerEUR.ReadToEnd();
                            Console.WriteLine(str);

                            int number1 = str.IndexOf(poiskEUR);
                            string kursEUR = str.Substring(number1 + 18, 6);


                            int numberDateEUR = str.IndexOf(poiskDateEUR);
                            string dateEUR = str.Substring(numberDateEUR + 7, 10);

                            DateTime dat = Convert.ToDateTime(dateEUR);
                            string YYY = dat.Year.ToString();
                            string MMM = dat.Month.ToString();
                            string DDD = dat.Day.ToString();
                            string datYMD = YYY + "-" + DDD + "-" + MMM;


                            ///----------------------------------------------------------------
                            string connectionString = "Provider = SQLOLEDB.1; Persist Security Info = True; User ID = sa; Initial Catalog = General; Data Source = galsrv\\ins1; Password = Cd3pk7zr% ";
                            using (OleDbConnection connection = new OleDbConnection(connectionString))
                            {
                                string comm1 = "select dbo.ToDate(f$DATVAL), F$SUMRUBL from t$cursval where f$DATVAL = dbo.ToAtlDate('" + datYMD + "') and f$KODVALUT = 0x8001000000000005";
                                string comm2 = "insert into t$cursval(f$KODVALUT, f$DATVAL, F$SUMRUBL, F$CMAIN) values(0x8001000000000005, dbo.ToAtlDate('" + datYMD + "'), " + kursEUR + ", 0x8001000000000005)";
                                
                                OleDbCommand command1 = new OleDbCommand(comm1);
                                OleDbCommand command2 = new OleDbCommand(comm2);

                                command1.Connection = connection;
                                command2.Connection = connection;

                                // Open the connection and execute the insert command.
                                try
                                {
                                    connection.Open();
                                    command1.ExecuteNonQuery();

                                    OleDbDataReader reader1 = command1.ExecuteReader();
                                    if (reader1.HasRows) // если есть данные
                                    {
                                        using (StreamWriter sw = new StreamWriter(writePath, true, System.Text.Encoding.Default))
                                        {
                                            sw.Write(DateTime.Now);
                                            sw.WriteLine(" Обнаружен в базе курс Евро: ");

                                            sw.WriteLine("Дата\t\tКурс");

                                            while (reader1.Read()) // построчно считываем данные
                                            {
                                                object dats = reader1.GetValue(0);
                                                object kurs = reader1.GetValue(1);

                                                sw.WriteLine("{0} \t{1}", dats, kurs);

                                            }
                                            sw.WriteLine("--------------------------------------------------------------------------------------------------------------");
                                        }
                                    }

                                    else
                                    {
                                        command2.ExecuteNonQuery();
                                        connection.Close();
                                        using (StreamWriter sw = new StreamWriter(writePath, true, System.Text.Encoding.Default))
                                        {
                                            sw.Write(DateTime.Now);
                                            sw.Write(" Введен курс ЕВРО: ");
                                            sw.Write(kursEUR + " на дату: ");
                                            sw.WriteLine(dateEUR);
                                        }
                                    }

                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine(ex.Message);
                                    using (StreamWriter sw = new StreamWriter(writePath, true, System.Text.Encoding.Default))
                                    {
                                        sw.WriteLine("Ошибка");
                                        sw.Write(DateTime.Now + ": ");
                                        sw.WriteLine(ex.Message);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                using (StreamWriter sw = new StreamWriter(writePath, true, System.Text.Encoding.Default))
                {
                    sw.WriteLine("Ошибка");
                    sw.Write(DateTime.Now + ": ");
                    sw.WriteLine(ex.Message);
                }
            }

            //---------------------------------UAH----------------------------------------
            try
            {
                HttpWebRequest requestUAH = (HttpWebRequest)WebRequest.Create("https://www.nbrb.by/api/exrates/rates/UAH?parammode=2&ondate=" + datYMDin);

                //string writePath = @"D:\kurs.txt";
                string poiskUAH = "Cur_OfficialRate";
                string poiskDateUAH = "Date";

                using (HttpWebResponse responseUAH = (HttpWebResponse)requestUAH.GetResponse())
                {
                    using (Stream streamUAH = responseUAH.GetResponseStream())
                    {
                        using (StreamReader readerUAH = new StreamReader(streamUAH))
                        {
                            string str = readerUAH.ReadToEnd();
                            Console.WriteLine(str);

                            int number1 = str.IndexOf(poiskUAH);
                            string kursUAH = str.Substring(number1 + 18, 6);
                            decimal kurdec = Convert.ToDecimal(kursUAH.Replace(".", ","))/100;
                            string kur = Convert.ToString(kurdec).Replace(",", ".");

                            int numberDateUAH = str.IndexOf(poiskDateUAH);
                            string dateUAH = str.Substring(numberDateUAH + 7, 10);

                            DateTime dat = Convert.ToDateTime(dateUAH);
                            string YYY = dat.Year.ToString();
                            string MMM = dat.Month.ToString();
                            string DDD = dat.Day.ToString();
                            string datYMD = YYY + "-" + DDD + "-" + MMM;


                            ///----------------------------------------------------------------
                            string connectionString = "Provider = SQLOLEDB.1; Persist Security Info = True; User ID = sa; Initial Catalog = General; Data Source = galsrv\\ins1; Password = Cd3pk7zr% ";
                            using (OleDbConnection connection = new OleDbConnection(connectionString))
                            {
                                string comm1 = "select dbo.ToDate(f$DATVAL), F$SUMRUBL from t$cursval where f$DATVAL = dbo.ToAtlDate('" + datYMD + "') and f$KODVALUT = 0x8001000000000004";
                                string comm2 = "insert into t$cursval(f$KODVALUT, f$DATVAL, F$SUMRUBL, F$CMAIN) values(0x8001000000000004, dbo.ToAtlDate('" + datYMD + "'), " + kur + ", 0x8001000000000004)";
                                
                                OleDbCommand command1 = new OleDbCommand(comm1);
                                OleDbCommand command2 = new OleDbCommand(comm2);

                                command1.Connection = connection;
                                command2.Connection = connection;

                                // Open the connection and execute the insert command.
                                try
                                {
                                    connection.Open();
                                    command1.ExecuteNonQuery();

                                    OleDbDataReader reader1 = command1.ExecuteReader();
                                    if (reader1.HasRows) // если есть данные
                                    {
                                        using (StreamWriter sw = new StreamWriter(writePath, true, System.Text.Encoding.Default))
                                        {
                                            sw.Write(DateTime.Now);
                                            sw.WriteLine(" Обнаружен в базе курс Гривни: ");

                                            sw.WriteLine("Дата\t\tКурс");

                                            while (reader1.Read()) // построчно считываем данные
                                            {
                                                object dats = reader1.GetValue(0);
                                                object kurs = reader1.GetValue(1);

                                                sw.WriteLine("{0} \t{1}", dats, kurs);

                                            }
                                            sw.WriteLine("--------------------------------------------------------------------------------------------------------------");
                                        }
                                    }

                                    else
                                    {
                                        command2.ExecuteNonQuery();
                                        connection.Close();
                                        using (StreamWriter sw = new StreamWriter(writePath, true, System.Text.Encoding.Default))
                                        {
                                            sw.Write(DateTime.Now);
                                            sw.Write(" Введен курс Гривни: ");
                                            sw.Write(kursUAH + " на дату: ");
                                            sw.WriteLine(dateUAH);
                                        }
                                    }

                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine(ex.Message);
                                    using (StreamWriter sw = new StreamWriter(writePath, true, System.Text.Encoding.Default))
                                    {
                                        sw.WriteLine("Ошибка");
                                        sw.Write(DateTime.Now + ": ");
                                        sw.WriteLine(ex.Message);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                using (StreamWriter sw = new StreamWriter(writePath, true, System.Text.Encoding.Default))
                {
                    sw.WriteLine("Ошибка");
                    sw.Write(DateTime.Now + ": ");
                    sw.WriteLine(ex.Message);
                }
            }

            //---------------------RUB-----------------------------------------------------
            try
            {
                HttpWebRequest requestRUB = (HttpWebRequest)WebRequest.Create("https://www.nbrb.by/api/exrates/rates/RUB?parammode=2&ondate=" + datYMDin);

                string poiskRUB = "Cur_OfficialRate";
                string poiskDateRUB = "Date";

                using (HttpWebResponse responseRUB = (HttpWebResponse)requestRUB.GetResponse())
                {
                    using (Stream streamRUB = responseRUB.GetResponseStream())
                    {
                        using (StreamReader readerRUB = new StreamReader(streamRUB))
                        {
                            string str = readerRUB.ReadToEnd();
                            Console.WriteLine(str);

                            int number1 = str.IndexOf(poiskRUB);
                            string kursRUB = str.Substring(number1 + 18, 6);


                            int numberDateRUB = str.IndexOf(poiskDateRUB);
                            string dateRUB = str.Substring(numberDateRUB + 7, 10);

                            DateTime dat = Convert.ToDateTime(dateRUB);
                            string YYY = dat.Year.ToString();
                            string MMM = dat.Month.ToString();
                            string DDD = dat.Day.ToString();
                            string datYMD = YYY + "-" + DDD + "-" + MMM;

                            ///----------------------------------------------------------------
                            string connectionString = "Provider = SQLOLEDB.1; Persist Security Info = True; User ID = sa; Initial Catalog = General; Data Source = galsrv\\ins1; Password = Cd3pk7zr% ";
                            using (OleDbConnection connection = new OleDbConnection(connectionString))
                            {
                                string comm1 = "select dbo.ToDate(f$DATVAL), F$SUMRUBL from t$cursval where f$DATVAL = dbo.ToAtlDate('" + datYMD + "') and f$KODVALUT = 0x8000000000000004";
                                string comm2 = "insert into t$cursval(f$KODVALUT, f$DATVAL, F$SUMRUBL, F$CMAIN) values(0x8000000000000004, dbo.ToAtlDate('" + datYMD + "'), " + kursRUB + ", 0x8000000000000004)";

                                OleDbCommand command1 = new OleDbCommand(comm1);
                                OleDbCommand command2 = new OleDbCommand(comm2);

                                command1.Connection = connection;
                                command2.Connection = connection;

                                // Open the connection and execute the insert command.
                                try
                                {
                                    connection.Open();
                                    command1.ExecuteNonQuery();

                                    OleDbDataReader reader1 = command1.ExecuteReader();
                                    if (reader1.HasRows) // если есть данные
                                    {
                                        using (StreamWriter sw = new StreamWriter(writePath, true, System.Text.Encoding.Default))
                                        {
                                            sw.Write(DateTime.Now);
                                            sw.WriteLine(" Обнаружен в базе курс Российского рубля: ");

                                            sw.WriteLine("Дата\t\tКурс");

                                            while (reader1.Read()) // построчно считываем данные
                                            {
                                                object dats = reader1.GetValue(0);
                                                object kurs = reader1.GetValue(1);

                                                sw.WriteLine("{0} \t{1}", dats, kurs);

                                            }
                                            sw.WriteLine("--------------------------------------------------------------------------------------------------------------");
                                        }
                                    }

                                    else
                                    {
                                        command2.ExecuteNonQuery();
                                        connection.Close();
                                        using (StreamWriter sw = new StreamWriter(writePath, true, System.Text.Encoding.Default))
                                        {
                                            sw.Write(DateTime.Now);
                                            sw.Write(" Введен курс Российского рубля: ");
                                            sw.Write(kursRUB + " на дату: ");
                                            sw.WriteLine(dateRUB);
                                        }
                                    }

                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine(ex.Message);
                                    using (StreamWriter sw = new StreamWriter(writePath, true, System.Text.Encoding.Default))
                                    {
                                        sw.WriteLine("Ошибка");
                                        sw.Write(DateTime.Now + ": ");
                                        sw.WriteLine(ex.Message);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                using (StreamWriter sw = new StreamWriter(writePath, true, System.Text.Encoding.Default))
                {
                    sw.WriteLine("Ошибка");
                    sw.Write(DateTime.Now + ": ");
                    sw.WriteLine(ex.Message);
                }
            }
            //---------------------------------------------------------------------------
            using (StreamWriter sw = new StreamWriter(writePath, true, System.Text.Encoding.Default))
            {
                sw.WriteLine("--------------------------------------------------------------------------------------------------------------");
            }
        }
    }
}
