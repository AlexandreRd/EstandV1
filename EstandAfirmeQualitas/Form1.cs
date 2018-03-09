using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
// Access DB
using System.Data.OleDb;
// UI Thread 
using System.Threading;
// CSV Read
using System.IO;
// Data Analysis
using RDotNet;
// Levenshtein Distance
using MinimumEditDistance;
// JaroWinkler Distance
//using FuzzyString;
using SimMetricsMetricUtilities;

namespace EstandAfirmeQualitas
{
    public partial class Form1 : Form
    {
        // Declare our worker thread
        private Thread workerThread = null;
        // Declare a delegate used to communicate with the UI thread
        private delegate void UpdateStatusDelegate();
        private UpdateStatusDelegate updateStatusDelegate = null;
        
        // Homologation Count
        int pBMax = 1;
        int pBCount = 0;
        DateTime timeStart;
        TimeSpan timeExec;

        // Connection to DB
        String MyConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Users\\Griselda\\Documents\\Nueva BD\\Homologacion.mdb;";

        private void UpdateStatus()
        {
            timeExec = DateTime.Now - timeStart;
            txtExec.Text = String.Format("{0:D2}:{1:D2}:{2:D2}", timeExec.Hours, timeExec.Minutes, timeExec.Seconds);
            txtCount.Text = pBCount.ToString() + " of " + pBMax.ToString(); 
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.updateStatusDelegate = new UpdateStatusDelegate(this.UpdateStatus);
            //this.setBarDelegate = new SetBarDelegate(() => this.setPbar(pBMax));
        }

        public Form1()
        {
            InitializeComponent();
        }

        // Get Acronym Database from .CSV
        public List<String>[] getAcroDB()
        {
            using (var readCSV = new StreamReader(@"C:\Users\Griselda\Documents\AcronymDB.csv"))
            {
                List<String>[] acrDB = new List<String>[12];
                for (Int32 i = 0; i < acrDB.Length; i++) {
                    acrDB[i] = new List<String>();    
                }

                int reg = 0;
                while (!readCSV.EndOfStream)
                {
                    var line = readCSV.ReadLine();
                    var values = line.Split(',');
                    
                    if (reg > 0)
                    {
                        acrDB[0].Add((values[0].ToString() != "") ? values[0]: "NA");
                        // MessageBox.Show(acrDB[0].ElementAt(acrDB[0].Count-1).ToString(), "Trans");
                        acrDB[1].Add((values[1].ToString() != "") ? values[1] : "NA");
                        acrDB[2].Add((values[2].ToString() != "") ? values[2] : "NA");
                        acrDB[3].Add((values[3].ToString() != "") ? values[3] : "NA");
                        acrDB[4].Add((values[4].ToString() != "") ? values[4] : "NA");
                        acrDB[5].Add((values[5].ToString() != "") ? values[5] : "NA");
                        acrDB[6].Add((values[6].ToString() != "") ? values[6] : "NA");
                        acrDB[7].Add((values[7].ToString() != "") ? values[7] : "NA");
                        acrDB[8].Add((values[8].ToString() != "") ? values[8] : "NA");
                        acrDB[9].Add((values[9].ToString() != "") ? values[9] : "NA");
                        acrDB[10].Add((values[10].ToString() != "") ? values[10] : "NA");
                        acrDB[11].Add((values[11].ToString() != "") ? values[11] : "NA");
                    }
                    reg++;
                }
                
                MessageBox.Show("Acronym Database Loaded","SUCCESS");
                return acrDB;
            }
        } 

        // VERIFICAR MATCH EN DATATABLE
        public Int32 evalReader(DataTable myStandData) 
        {
            // JaroWinkler Object
            var Jw = new JaroWinkler();

            int RowC = 0;
            
            double[] MatchD = new double[myStandData.Rows.Count];
            Int32[] MatchF = new Int32[myStandData.Rows.Count];
            
            DataRow MyModel = null;

            // Evaluando Elementos
            foreach (DataRow Row in myStandData.Rows) 
            {
                if (RowC == 0)
                {
                    MyModel = Row;
                }
                else
                {
                    // Transmision
                    if (equalDescrip(MyModel.Field<String>(0), Row.Field<String>(0)) || Row.Field<String>(0).Length == 0)
                    {
                        MatchF[RowC]++;
                    }
                    // GearBox
                    if (equalDescrip(MyModel.Field<String>(1), Row.Field<String>(1)) || Row.Field<String>(1).Length == 0)
                    {
                        MatchF[RowC]++;
                    }
                    // Cilindros
                    if (equalDescrip(MyModel.Field<String>(2), Row.Field<String>(2)) || Row.Field<String>(2).Length == 0)
                    {
                        MatchF[RowC]++;
                    }
                    // Pasajeros
                    if (equalDescrip(MyModel.Field<String>(3), Row.Field<String>(3)) || Row.Field<String>(3).Length == 0)
                    {
                        MatchF[RowC]++;
                    }
                    // Puertas
                    if (equalDescrip(MyModel.Field<String>(4), Row.Field<String>(4)) || Row.Field<String>(4).Length == 0)
                    {
                        MatchF[RowC]++;
                    }
                    // ABS
                    if (equalDescrip(MyModel.Field<String>(5), Row.Field<String>(5)) || Row.Field<String>(5).Length == 0)
                    {
                        MatchF[RowC]++;
                    }
                    // Vestidura
                    if (equalDescrip(MyModel.Field<String>(6), Row.Field<String>(6)) || Row.Field<String>(6).Length == 0)
                    {
                        MatchF[RowC]++;
                    }
                    // Sonido
                    if (equalDescrip(MyModel.Field<String>(7), Row.Field<String>(7)) || Row.Field<String>(7).Length == 0)
                    {
                        MatchF[RowC]++;
                    }
                    // Equipado
                    if (equalDescrip(MyModel.Field<String>(8), Row.Field<String>(8)) || Row.Field<String>(8).Length == 0)
                    {
                        MatchF[RowC]++;
                    }
                    // Aire
                    if (equalDescrip(MyModel.Field<String>(9), Row.Field<String>(9)) || Row.Field<String>(9).Length == 0)
                    {
                        MatchF[RowC]++;
                    }
                    // BAire
                    if (equalDescrip(MyModel.Field<String>(10), Row.Field<String>(10)) || Row.Field<String>(10).Length == 0)
                    {
                        MatchF[RowC]++;
                    }
                    // QC
                    if (equalDescrip(MyModel.Field<String>(11), Row.Field<String>(11)) || Row.Field<String>(11).Length == 0)
                    {
                        MatchF[RowC]++;
                    }
                    // DescripSimple
                    MatchD[RowC] = (Row.Field<String>(12).Length > 0) ? 
                        Jw.GetSimilarity(sortDescrip(MyModel.Field<String>(12)), sortDescrip(Row.Field<String>(12))) 
                        : 0.70;
                    /*
                    if (Levenshtein.CalculateDistance(
                            sortDescrip(MyModel.Field<String>(12)), sortDescrip(Row.Field<String>(12)), 1)
                            < Convert.ToInt32(Math.Max(Row.Field<String>(12).Length, MyModel.Field<String>(12).Length) * 0.7))
                    {
                             MatchD[RowC] = true;
                    }  */
                }
                RowC++;
            }

            int bestMatch = 0;
            for (int i = 0; i < MatchF.Length; i++) {
                if (i == 0) {
                    MatchD[i] = 0.699;
                    MatchF[i] = 6;
                }
                else
                {
                    if (MatchF[i] > MatchF[bestMatch] && MatchD[i] > MatchD[bestMatch])
                        bestMatch = i;
                }
            }

            return bestMatch = (MatchF[bestMatch] > 7 && MatchD[bestMatch] > 0.7) ? bestMatch : 0;
        }

        // ORDENAR STRING ALFABETICAMENTE
        public String sortDescrip(String a)
        {
            a = a.ToUpper();
            String aWord = "";
            List<String> listA = new List<String>();
            int aux = 0;

            // Descomponer el String A en palabras y agregarlo a List
            while (a.Length > 0)
            {
                // Remover ' ' al inicio 
                a = a.Trim();

                // String valido
                if (a.Length > 0)
                {
                    // Varias palabras
                    if (a.LastIndexOf(' ') > 0)
                    {
                        aux = a.IndexOf(' ');
                        aWord = a.Substring(0, aux);
                        // MessageBox.Show("_" + aWord + "_", "Word A");
                        a = a.Remove(0, aux);
                        // MessageBox.Show("_" + a + "_", "String restante A");
                    }
                    // Ultima palabra
                    else
                    {
                        aWord = a;
                        // MessageBox.Show("_" + aWord + "_", "Last Word A");
                        a = "";
                    }
                    // Agregando palabra a la lista
                    listA.Add(aWord);
                    // MessageBox.Show(listA.Count.ToString(), "Words in listA");
                    aWord = ""; aux = 0;
                }
                else
                {
                    break;
                }
            }

            listA.Sort();

            aWord = "";
            foreach (String Val in listA)
            {
                aWord += Val + " ";
            }

            return aWord;
        }

        // VERIFICAR QUE DOS CADENAS SEAN IGUALES
        public Boolean equalDescrip(String a, String b) {
            a = a.ToUpper(); b = b.ToUpper();
            String aWord = "", bWord = "";
            List<String> listA = new List<String>();
            List<String> listB = new List<String>();
            int aux = 0;

            // Descomponer el String A en palabras y agregarlo a List
            while (a.Length > 0) {
                // Remover ' ' al inicio 
                a = a.Trim();
                // MessageBox.Show("_" + a + "_", "String A sin Esp");

                // String valido
                if (a.Length > 0) {
                    // Varias palabras
                    if (a.LastIndexOf(' ') > 0) {
                        aux = a.IndexOf(' ');
                        aWord = a.Substring(0, aux);
                        // MessageBox.Show("_" + aWord + "_", "Word A");
                        a = a.Remove(0, aux);
                        // MessageBox.Show("_" + a + "_", "String restante A");
                    }
                    // Ultima palabra
                    else {
                        aWord = a;
                        // MessageBox.Show("_" + aWord + "_", "Last Word A");
                        a = "";
                    }
                    // Agregando palabra a la lista
                    listA.Add(aWord);
                    // MessageBox.Show(listA.Count.ToString(), "Words in listA");
                    aWord = ""; aux = 0;
                } else {
                    break;
                }
            }

            // Descomponer el String B en palabras y agregarlo a List
            while (b.Length > 0) {
                // Remover ' ' al inicio 
                b = b.Trim();

                // String valido
                if (b.Length > 0) {
                    // Varias palabras
                    if (b.LastIndexOf(' ') > 0) {
                        aux = b.IndexOf(' ');
                        bWord = b.Substring(0, aux);
                        b = b.Remove(0, aux);
                    }
                    // Ultima palabra
                    else {
                        bWord = b;
                        b = "";
                    }
                    // Agregando palabra a la lista
                    listB.Add(bWord);
                    bWord = ""; aux = 0;
                }
                else {
                    break;
                }
            }

            // Verificar que las listas tienen el mismo tamaño y contienen los mismos elementos
            return (listA.Count == listB.Count) && new HashSet<string>(listA).SetEquals(listB);                  
        }

        // EJECUTAR CONSULTA
        public void doQuery(String CONS, String CONEX)
        {
            OleDbConnection CONNECT = new OleDbConnection(CONEX
                //"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Users\\Griselda\\Documents\\Nueva BD\\Test.mdb;"
                );
            try
            {
                CONNECT.Open();
                OleDbCommand COMMAND = new OleDbCommand(CONS, CONNECT);
                OleDbDataReader READER = COMMAND.ExecuteReader();
                READER.Close();
                CONNECT.Close();
                // MessageBox.Show("SUCCESSS");
            }
            catch (Exception e)
            {
                MessageBox.Show("QUERY: ___" + CONS + "___ EX:___" + e.ToString(), "ERROR IN QUERY");
            }

        }

        // GENERAR TABLAS ESTANDARIZADAS
        // QUALITAS
        public void qualitasToStand()
        {
            pBCount = 0;
            try
            {
                OleDbConnection CONNECTPB = new OleDbConnection(MyConnString);
                CONNECTPB.Open();
                OleDbCommand COMMAND_PBAR = new OleDbCommand("SELECT COUNT (*) FROM Tar_Qualitas_2017", CONNECTPB);

                pBMax = Convert.ToInt32(COMMAND_PBAR.ExecuteScalar());

                CONNECTPB.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.ToString(), "Error en Count de Registros");
            }

            OleDbConnection CONNECT = new OleDbConnection(
                MyConnString);
            // C:\Users\Griselda\Desktop\TarifasAlex
            String CONSULT_MOD = "SELECT * FROM Tar_Qualitas_2017";
            String CONSULT_HOM = "SELECT * FROM D_Qualitas WHERE ";
            try
            {
                CONNECT.Open();
                OleDbCommand COMMAND_MOD = new OleDbCommand(CONSULT_MOD, CONNECT);
                OleDbDataReader READER_MOD = COMMAND_MOD.ExecuteReader();

                while (READER_MOD.Read())
                {
                    OleDbCommand COMMAND_HOM = new OleDbCommand(CONSULT_HOM +
                       "Clave = '" + Convert.ToInt32(READER_MOD["CAMIS"]).ToString() + "'",
                       CONNECT);
                    OleDbDataReader READER_HOM = COMMAND_HOM.ExecuteReader();
                    while (READER_HOM.Read())
                    {
                        String desTSM = (READER_HOM["DescripSimple"].ToString().Trim().Length > 0) ?
                            READER_HOM["DescripSimple"].ToString().Trim() + " " : "";
                        desTSM += (READER_HOM["Equipado"].ToString().Trim().Length > 0) ? READER_HOM["Equipado"].ToString().Trim() + " " : "";
                        desTSM += (READER_HOM["Trans"].ToString().Trim().Length > 0) ? READER_HOM["Trans"].ToString().Trim() + " " : "";
                        desTSM += (READER_HOM["Puertas"].ToString().Trim().Length > 0) ? READER_HOM["Puertas"].ToString().Trim() + " " : "";
                        desTSM += (READER_HOM["Vestiduras"].ToString().Trim().Length > 0) ? READER_HOM["Vestiduras"].ToString().Trim() + " " : "";
                        desTSM += (READER_HOM["ABS"].ToString().Trim().Length > 0) ? READER_HOM["ABS"].ToString().Trim() + " " : "";
                        desTSM += (READER_HOM["Aire"].ToString().Trim().Length > 0) ? READER_HOM["Aire"].ToString().Trim() + " " : "";
                        desTSM += (READER_HOM["QC"].ToString().Trim().Length > 0) ? READER_HOM["QC"].ToString().Trim() + " " : "";
                        desTSM += (Convert.ToInt32(READER_HOM["NPass"].ToString()) > 0) ? READER_HOM["NPass"].ToString().Trim() + "Pasaj " : "";
                        desTSM += (READER_HOM["EE"].ToString().Trim().Length > 0) ? READER_HOM["EE"].ToString().Trim() + " " : "";
                        desTSM += (READER_HOM["Vidrios"].ToString().Trim().Length > 0) ? READER_HOM["Vidrios"].ToString().Trim() + " " : "";
                        desTSM += (READER_HOM["BAire"].ToString().Trim().Length > 0) ? READER_HOM["BAire"].ToString().Trim() + " " : "";
                        desTSM += (READER_HOM["Sonido"].ToString().Trim().Length > 0) ? READER_HOM["Sonido"].ToString().Trim() + " " : ""; ;
                        desTSM += (READER_HOM["FN"].ToString().Trim().Length > 0) ? READER_HOM["FN"].ToString().Trim() + " " : "";
                        desTSM += (READER_HOM["DH"].ToString().Trim().Length > 0) ? READER_HOM["DH"].ToString().Trim() + " " : "";
                        desTSM += (READER_HOM["DT"].ToString().Trim().Length > 0) ? READER_HOM["DT"].ToString().Trim() + " " : "";
                        desTSM += (READER_HOM["RA"].ToString().Trim().Length > 0) ? READER_HOM["RA"].ToString().Trim() + " " : "";
                        desTSM += (READER_HOM["Cilindros"].ToString().Trim().Length > 0) ? READER_HOM["Cilindros"].ToString().Trim() + " " : "";
                        desTSM += (READER_HOM["FL"].ToString().Trim().Length > 0) ? READER_HOM["FL"].ToString().Trim() + " " : "";
                        desTSM += (READER_HOM["BF"].ToString().Trim().Length > 0) ? READER_HOM["BF"].ToString().Trim() + " " : "";
                        desTSM += (READER_HOM["PE"].ToString().Trim().Length > 0) ? READER_HOM["PE"].ToString().Trim() + " " : "";
                        desTSM += (READER_HOM["TP"].ToString().Trim().Length > 0) ? READER_HOM["TP"].ToString().Trim() + " " : "";
                        desTSM += (READER_HOM["TC"].ToString().Trim().Length > 0) ? READER_HOM["TC"].ToString().Trim() + " " : "";
                        desTSM += (READER_HOM["CodRaro"].ToString().Trim().Length > 0) ? READER_HOM["CodRaro"].ToString().Trim() + " " : "";

                        doQuery("INSERT INTO Estandarizados_Qualitas" +
                                "(Cia, TipoTar, Clave, Marca, Tipo, Modelo, DescripCia, DescripSimple, " +
                                "Equipado, Trans, Puertas, Vestiduras, ABS, Aire, QC, NPass, EE, Vidrios, BAire, Sonido, " +
                                "FN, DH, DT, RA, Cilindros, FL, BF, PE, TP, TC, CodRaro, " +
                                "DescripTSM)" +
                                "VALUES (" +
                                "2, 0, " +
                                READER_HOM["Clave"].ToString().Trim() + ", '" +
                                READER_HOM["Marca"].ToString().Trim() + "', '" +
                                READER_HOM["Tipo"].ToString().Trim() + "', " +

                                READER_MOD["cModelo"].ToString().Trim() + ", '" +

                                READER_HOM["DescripCia"].ToString().Trim() + "', '" +
                                READER_HOM["DescripSimple"].ToString().Trim() + "', '" +
                                READER_HOM["Equipado"].ToString().Trim() + "', '" +
                                READER_HOM["Trans"].ToString().Trim() + "', '" +
                                READER_HOM["Puertas"].ToString().Trim() + "', '" +
                                READER_HOM["Vestiduras"].ToString().Trim() + "', '" +
                                READER_HOM["ABS"].ToString().Trim() + "', '" +
                                READER_HOM["Aire"].ToString().Trim() + "', '" +
                                READER_HOM["QC"].ToString().Trim() + "', '" +
                                ((Convert.ToInt32(READER_HOM["NPass"].ToString()) > 0) ? READER_HOM["NPass"].ToString() + "Pasaj" : "") + "', '" +
                                READER_HOM["EE"].ToString().Trim() + "', '" +
                                READER_HOM["Vidrios"].ToString().Trim() + "', '" +
                                READER_HOM["BAire"].ToString().Trim() + "', '" +
                                READER_HOM["Sonido"].ToString().Trim() + "', '" +
                                READER_HOM["FN"].ToString().Trim() + "', '" +
                                READER_HOM["DH"].ToString().Trim() + "', '" +
                                READER_HOM["DT"].ToString().Trim() + "', '" +
                                READER_HOM["RA"].ToString().Trim() + "', '" +
                                READER_HOM["Cilindros"].ToString().Trim() + "', '" +
                                READER_HOM["FL"].ToString().Trim() + "', '" +
                                READER_HOM["BF"].ToString().Trim() + "', '" +
                                READER_HOM["PE"].ToString().Trim() + "', '" +
                                READER_HOM["TP"].ToString().Trim() + "', '" +
                                READER_HOM["TC"].ToString().Trim() + "', '" +
                                READER_HOM["CodRaro"].ToString().Trim() + "', '" +
                                desTSM +
                                "')",

                                MyConnString
                                );

                        //MessageBox.Show("NUEVO REGISTRO", "INFO");
                    }
                    pBCount++;
                    this.Invoke(this.updateStatusDelegate);

                    READER_HOM.Close();
                }
                READER_MOD.Close();
                CONNECT.Close();
                MessageBox.Show("SUCCESS");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString(), "ERROR IN QUALITAS TO STAND");
            }
            finally {
                this.workerThread.Abort();
            }
        }

        // AFIRME
        public void afirmeToStand()
        {
            pBCount = 0;
            try
            {
                OleDbConnection CONNECTPB = new OleDbConnection(MyConnString);
                CONNECTPB.Open();
                OleDbCommand COMMAND_PBAR = new OleDbCommand("SELECT COUNT (*) FROM D_Afirme", CONNECTPB);

                pBMax = Convert.ToInt32(COMMAND_PBAR.ExecuteScalar());

                CONNECTPB.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.ToString(), "Error en Count de Registros");
            }

            OleDbConnection CONNECT = new OleDbConnection(
                MyConnString);
            String CONSULTA = "SELECT * FROM D_Afirme";
            try
            {
                CONNECT.Open();
                OleDbCommand COMMAND = new OleDbCommand(CONSULTA, CONNECT);
                OleDbDataReader READER = COMMAND.ExecuteReader();

                while (READER.Read())
                {
                    int iMod;
                    for (iMod = Convert.ToInt32(READER["ModeloI"]); iMod <= Convert.ToInt32(READER["ModeloF"]); iMod++)
                    {
                        String desTSM = (READER["DescripSimple"].ToString().Trim().Length > 0) ?
                            READER["DescripSimple"].ToString().Trim() + " " : "";
                        desTSM += (READER["Equipado"].ToString().Trim().Length > 0) ? READER["Equipado"].ToString().Trim() + " " : "";
                        desTSM += (READER["Trans"].ToString().Trim().Length > 0) ? READER["Trans"].ToString().Trim() + " " : "";
                        desTSM += (READER["Puertas"].ToString().Trim().Length > 0) ? READER["Puertas"].ToString().Trim() + " " : "";
                        desTSM += (READER["Vestiduras"].ToString().Trim().Length > 0) ? READER["Vestiduras"].ToString().Trim() + " " : "";
                        desTSM += (READER["ABS"].ToString().Trim().Length > 0) ? READER["ABS"].ToString().Trim() + " " : "";
                        desTSM += (READER["Aire"].ToString().Trim().Length > 0) ? READER["Aire"].ToString().Trim() + " " : "";
                        desTSM += (READER["QC"].ToString().Trim().Length > 0) ? READER["QC"].ToString().Trim() + " " : "";
                        desTSM += (Convert.ToInt32(READER["NPass"].ToString()) > 0) ? READER["NPass"].ToString().Trim() + "Pasaj " : "";
                        desTSM += (READER["EE"].ToString().Trim().Length > 0) ? READER["EE"].ToString().Trim() + " " : "";
                        desTSM += (READER["Vidrios"].ToString().Trim().Length > 0) ? READER["Vidrios"].ToString().Trim() + " " : "";
                        desTSM += (READER["BAire"].ToString().Trim().Length > 0) ? READER["BAire"].ToString().Trim() + " " : "";
                        desTSM += (READER["Sonido"].ToString().Trim().Length > 0) ? READER["Sonido"].ToString().Trim() + " " : ""; ;
                        desTSM += (READER["FN"].ToString().Trim().Length > 0) ? READER["FN"].ToString().Trim() + " " : "";
                        desTSM += (READER["DH"].ToString().Trim().Length > 0) ? READER["DH"].ToString().Trim() + " " : "";
                        desTSM += (READER["DT"].ToString().Trim().Length > 0) ? READER["DT"].ToString().Trim() + " " : "";
                        desTSM += (READER["RA"].ToString().Trim().Length > 0) ? READER["RA"].ToString() + " " : "";
                        desTSM += (READER["Cilindros"].ToString().Trim().Length > 0) ? READER["Cilindros"].ToString().Trim() + " " : "";
                        desTSM += (READER["FL"].ToString().Trim().Length > 0) ? READER["FL"].ToString().Trim() + " " : "";
                        desTSM += (READER["BF"].ToString().Trim().Length > 0) ? READER["BF"].ToString().Trim() + " " : "";
                        desTSM += (READER["PE"].ToString().Trim().Length > 0) ? READER["PE"].ToString().Trim() + " " : "";
                        desTSM += (READER["TP"].ToString().Trim().Length > 0) ? READER["TP"].ToString().Trim() + " " : "";
                        desTSM += (READER["TC"].ToString().Trim().Length > 0) ? READER["TC"].ToString().Trim() + " " : "";
                        desTSM += (READER["CodRaro"].ToString().Trim().Length > 0) ? READER["CodRaro"].ToString().Trim() + " " : "";

                        doQuery("INSERT INTO Estandarizados_Afirme" +
                               "(Cia, TipoTar, Clave, Marca, Tipo, Modelo, DescripCia, DescripSimple, " +
                               "Equipado, Trans, Puertas, Vestiduras, ABS, Aire, QC, NPass, EE, Vidrios, BAire, Sonido, " + 
                               "FN, DH, DT, RA, Cilindros, FL, BF, PE, TP, TC, CodRaro, " +
                                "DescripTSM)" +
                                "VALUES (" +
                                   "21, 0, " +
                               READER["Clave"].ToString().Trim() + ", '" +
                               READER["Marca"].ToString().Trim() + "', '" +
                               READER["Tipo"].ToString().Trim() + "', " +
                               iMod.ToString().Trim() + ", '" +
                               READER["DescripCia"].ToString().Trim() + "', '" +
                               READER["DescripSimple"].ToString().Trim() + "', '" +
                               READER["Equipado"].ToString().Trim() + "', '" +
                               READER["Trans"].ToString().Trim() + "', '" +
                               READER["Puertas"].ToString().Trim() + "', '" +
                               READER["Vestiduras"].ToString().Trim() + "', '" +
                               READER["ABS"].ToString().Trim() + "', '" +
                               READER["Aire"].ToString().Trim() + "', '" +
                               READER["QC"].ToString().Trim() + "', '" +
                               ((Convert.ToInt32(READER["NPass"].ToString()) > 0) ? READER["NPass"].ToString().Trim() + "Pasaj" : "") + "', '" +
                               READER["EE"].ToString().Trim() + "', '" +
                               READER["Vidrios"].ToString().Trim() + "', '" +
                               READER["BAire"].ToString().Trim() + "', '" +
                               READER["Sonido"].ToString().Trim() + "', '" +
                               READER["FN"].ToString().Trim() + "', '" +
                               READER["DH"].ToString().Trim() + "', '" +
                               READER["DT"].ToString().Trim() + "', '" +
                               READER["RA"].ToString().Trim() + "', '" +
                               READER["Cilindros"].ToString().Trim() + "', '" +
                               READER["FL"].ToString().Trim() + "', '" +
                               READER["BF"].ToString().Trim() + "', '" +
                               READER["PE"].ToString().Trim() + "', '" +
                               READER["TP"].ToString().Trim() + "', '" +
                               READER["TC"].ToString().Trim() + "', '" +
                               READER["CodRaro"].ToString().Trim() + "', '" +
                               desTSM + 
                               "')"
                               ,
                               MyConnString
                           );
                    }
                    pBCount++;
                    this.Invoke(this.updateStatusDelegate);
                }
                READER.Close();
                CONNECT.Close();
                MessageBox.Show("SUCCESS");

            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString(), "ERROR IN AFIRME TO STAND");
            }
            finally
            {
                this.workerThread.Abort();
            }
        }

        // DB GENERICA
        public void DBToStand(String Company, Int32 numCompany)
        {
            pBCount = 0;
            try
            {
                OleDbConnection CONNECTPB = new OleDbConnection(MyConnString);
                CONNECTPB.Open();
                OleDbCommand COMMAND_PBAR = new OleDbCommand("SELECT COUNT (*) FROM D_" + Company, CONNECTPB);

                pBMax = Convert.ToInt32(COMMAND_PBAR.ExecuteScalar());

                CONNECTPB.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.ToString(), "Error en Count de Registros");
            }

            OleDbConnection CONNECT = new OleDbConnection(
                MyConnString);
            String CONSULTA = "SELECT * FROM D_" + Company;
            var np = "";
            try
            {
                CONNECT.Open();
                OleDbCommand COMMAND = new OleDbCommand(CONSULTA, CONNECT);
                OleDbDataReader READER = COMMAND.ExecuteReader();

                while (READER.Read())
                {
                    if (READER["Tipo"].ToString() != "" && READER["Marca"].ToString() != "ESPECIALES" && READER["Marca"].ToString() != "FRONTERIZO" && READER["Marca"].ToString() != "LEGALIZADO" && READER["Marca"].ToString() != "FRONTERIZO" && READER["Marca"].ToString() != "PLAN PISO" && READER["Marca"].ToString() != "CLASICO")
                    {
                        String desTSM = (READER["DescripSimple"].ToString().Trim().Length > 0) ? 
                            READER["DescripSimple"].ToString().Trim() + " " : "";
                        desTSM += (READER["Equipado"].ToString().Trim().Length > 0) ? READER["Equipado"].ToString().Trim() + " " : "";
                        desTSM += (READER["Trans"].ToString().Trim().Length > 0) ? READER["Trans"].ToString().Trim() + " " : "";
                        desTSM += (READER["Puertas"].ToString().Trim().Length > 0) ? READER["Puertas"].ToString().Trim() + " " : "";
                        desTSM += (READER["Vestiduras"].ToString().Trim().Length > 0) ? READER["Vestiduras"].ToString().Trim() + " " : "";
                        desTSM += (READER["ABS"].ToString().Trim().Length > 0) ? READER["ABS"].ToString().Trim() + " " : "";
                        desTSM += (READER["Aire"].ToString().Trim().Length > 0) ? READER["Aire"].ToString().Trim() + " " : "";
                        desTSM += (READER["QC"].ToString().Trim().Length > 0) ? READER["QC"].ToString().Trim() + " " : "";
                        np = READER["Clave"].ToString().Trim() + " : " + READER["Modelo"].ToString().Trim();
                        desTSM += (Convert.ToInt32(READER["NPass"].ToString()) > 0) ? READER["NPass"].ToString().Trim() + "Pasaj " : "";
                        desTSM += (READER["EE"].ToString().Trim().Length > 0) ? READER["EE"].ToString().Trim() + " " : "";
                        desTSM += (READER["Vidrios"].ToString().Trim().Length > 0) ? READER["Vidrios"].ToString().Trim() + " " : "";
                        desTSM += (READER["BAire"].ToString().Trim().Length > 0) ? READER["BAire"].ToString().Trim() + " " : "";
                        desTSM += (READER["Sonido"].ToString().Trim().Length > 0) ? READER["Sonido"].ToString().Trim() + " " : ""; ;
                        desTSM += (READER["FN"].ToString().Trim().Length > 0) ? READER["FN"].ToString().Trim() + " " : "";
                        desTSM += (READER["DH"].ToString().Trim().Length > 0) ? READER["DH"].ToString().Trim() + " " : "";
                        desTSM += (READER["DT"].ToString().Trim().Length > 0) ? READER["DT"].ToString().Trim() + " " : "";
                        desTSM += (READER["RA"].ToString().Trim().Length > 0) ? READER["RA"].ToString().Trim() + " " : "";
                        desTSM += (READER["Cilindros"].ToString().Trim().Length > 0) ? READER["Cilindros"].ToString().Trim() + " " : "";
                        desTSM += (READER["FL"].ToString().Trim().Length > 0) ? READER["FL"].ToString().Trim() + " " : "";
                        desTSM += (READER["BF"].ToString().Trim().Length > 0) ? READER["BF"].ToString().Trim() + " " : "";
                        desTSM += (READER["PE"].ToString().Trim().Length > 0) ? READER["PE"].ToString().Trim() + " " : "";
                        desTSM += (READER["TP"].ToString().Trim().Length > 0) ? READER["TP"].ToString().Trim() + " " : "";
                        desTSM += (READER["TC"].ToString().Trim().Length > 0) ? READER["TC"].ToString().Trim() + " " : "";
                        desTSM += (READER["CodRaro"].ToString().Trim().Length > 0) ? READER["CodRaro"].ToString().Trim() + " " : "";

                        doQuery("INSERT INTO Estandarizados_" + Company + 
                               "(Cia, TipoTar, Clave, Marca, Tipo, Modelo, DescripCia, DescripSimple, " + 
                               "Equipado, Trans, Puertas, Vestiduras, ABS, Aire, QC, NPass, EE, Vidrios, BAire, Sonido, " + 
                               "FN, DH, DT, RA, Cilindros, FL, BF, PE, TP, TC, CodRaro, "
                               + "DescripTSM)" +
                                "VALUES (" + 
                               numCompany.ToString() +
                                     ", 0, '" +
                               READER["Clave"].ToString().Trim() + "', '" +
                               READER["Marca"].ToString().Trim() + "', '" +
                               READER["Tipo"].ToString().Trim() + "', " +
                               READER["Modelo"].ToString() + ", '" +
                               READER["DescripCia"].ToString().Trim() + "', '" +
                               READER["DescripSimple"].ToString().Trim() + "', '" +
                               READER["Equipado"].ToString().Trim() + "', '" +
                               READER["Trans"].ToString().Trim() + "', '" +
                               READER["Puertas"].ToString().Trim() + "', '" +
                               READER["Vestiduras"].ToString().Trim() + "', '" +
                               READER["ABS"].ToString().Trim() + "', '" +
                               READER["Aire"].ToString().Trim() + "', '" +
                               READER["QC"].ToString().Trim() + "', '" +
                                    ((Convert.ToInt32(READER["NPass"].ToString()) > 0) ? READER["NPass"].ToString().Trim() + "Pasaj" : "") + "', '" +
                               READER["EE"].ToString().Trim() + "', '" +
                               READER["Vidrios"].ToString().Trim() + "', '" +
                               READER["BAire"].ToString().Trim() + "', '" +
                               READER["Sonido"].ToString().Trim() + "', '" +
                               READER["FN"].ToString().Trim() + "', '" +
                               READER["DH"].ToString().Trim() + "', '" +
                               READER["DT"].ToString().Trim() + "', '" +
                               READER["RA"].ToString().Trim() + "', '" +
                               READER["Cilindros"].ToString().Trim() + "', '" +
                               READER["FL"].ToString().Trim() + "', '" +
                               READER["BF"].ToString().Trim() + "', '" +
                               READER["PE"].ToString().Trim() + "', '" +
                               READER["TP"].ToString().Trim() + "', '" +
                               READER["TC"].ToString().Trim() + "', '" +
                               READER["CodRaro"].ToString().Trim() + "', '" +
                               desTSM.Trim() +
                               "')"
                               ,
                               MyConnString
                           );
                    }
                    pBCount++;
                    this.Invoke(this.updateStatusDelegate);
                }
                READER.Close();
                CONNECT.Close();
                MessageBox.Show("SUCCESS");

            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString(), "ERROR IN " + Company + " TO STAND" + np.ToString());
            }
            finally
            {
                this.workerThread.Abort();
            }
        }     

        // INSERTAR TABLAS ESTANDARIZADAS EN DATOS ESTANDARIZADOS
        // INSERTAR COMPAÑIA Y GENERAR CEVIC
        public void insertNewComp(int numCia, String nomCia)
        {
            OleDbConnection CONNECT = new OleDbConnection(MyConnString);
            OleDbConnection CONNECT2 = new OleDbConnection(MyConnString);
            // Cambiar nombre de acuerdo con la Tabla de la Cia
            String CONSULT_COMP = "SELECT * FROM Estandarizados_" + nomCia; // +"2";
            String CONSULT_NEW = "SELECT * FROM DatosEstandarizados WHERE ";

            // Campos de CEVIC
            String cveCEVIC, Mar, Typ, Mod, cveCo;
            int nMod = 0;

            // Acronym DataBase
            // Fields: Trans, Gear, Pts, Pass, Brakes, Vest, Sound, Equip, Air, AirBag, QC, Descrip
            List<String>[] AcroDB = getAcroDB();

            // DataTable from DatosEstandarizados
            DataTable dTStand = new DataTable();
            dTStand.Columns.Add("sTrans", typeof(String));
            dTStand.Columns.Add("sGear", typeof(String));
            dTStand.Columns.Add("sCyl", typeof(String));
            dTStand.Columns.Add("sPts", typeof(String));
            dTStand.Columns.Add("sPass", typeof(String));
            dTStand.Columns.Add("sBrakes", typeof(String));
            dTStand.Columns.Add("sVest", typeof(String));
            dTStand.Columns.Add("sSound", typeof(String));
            dTStand.Columns.Add("sEquip", typeof(String));
            dTStand.Columns.Add("sAir", typeof(String));
            dTStand.Columns.Add("sAirBag", typeof(String));
            dTStand.Columns.Add("sQC", typeof(String));
            dTStand.Columns.Add("sDescrip", typeof(String));
            
            List<String> CEVList = new List<String>();
            
            // Progress Count
            pBCount = 0;
            try
            {
                OleDbConnection CONNECTPB = new OleDbConnection(MyConnString);
                CONNECTPB.Open();
                OleDbCommand COMMAND_PBAR = new OleDbCommand("SELECT COUNT (*) FROM Estandarizados_" + nomCia, CONNECTPB);
                
                pBMax = Convert.ToInt32(COMMAND_PBAR.ExecuteScalar());

                CONNECTPB.Close();
            }
            catch (Exception Ex) {
                MessageBox.Show(Ex.ToString(), "Error en Count de Registros");
            }
            //this.Invoke(this.setBarDelegate);

            try
            {
                CONNECT.Open();
                OleDbCommand COMMAND_COMP = new OleDbCommand(CONSULT_COMP, CONNECT);
                OleDbDataReader READER_COMP = COMMAND_COMP.ExecuteReader();

                while (READER_COMP.Read())
                {
                    try
                    {
                        CONNECT2.Open();
                        OleDbCommand COMMAND_NEW = new OleDbCommand(CONSULT_NEW +
                            "CEVIC IN (" +
                                "SELECT MyCEVICPool FROM (" +
                                    "SELECT IIF([Cia_" + numCia.ToString() + "] Is Null, 'unavailable', [Cia_" +numCia.ToString() + "]) AS MyCIA, CEVIC AS MyCEVICPool FROM DatosEstandarizados " +
                                    "WHERE Marca = '" + READER_COMP["Marca"].ToString() + "' AND Tipo = '" + READER_COMP["Tipo"].ToString() + "' AND Modelo = '" + READER_COMP["Modelo"].ToString() + "')" + 
                                "WHERE MyCIA = 'unavailable')",
                            CONNECT2);
                        OleDbDataReader READER_NEW = COMMAND_NEW.ExecuteReader();

                        String myQuery = "";
                        // POSIBLE MATCH
                        if (READER_NEW.HasRows)
                        {
                            CEVList.Clear();
                            dTStand.Rows.Clear();
                          	// Adding New Model as First Row
                     		dTStand.Rows.Add(
                     				READER_COMP["Trans"].ToString(), 
                     				READER_COMP["Trans"].ToString(), 
                     				READER_COMP["Cilindros"].ToString(), 
                     				READER_COMP["Puertas"].ToString(), 
                     				READER_COMP["NPass"].ToString(), 
                     				READER_COMP["ABS"].ToString(), 
                     				READER_COMP["Vestiduras"].ToString(),
                     				READER_COMP["Sonido"].ToString(), 
                     				READER_COMP["Equipado"].ToString(), 
                     				READER_COMP["Aire"].ToString(), 
                     				READER_COMP["BAire"].ToString(), 
                     				READER_COMP["QC"].ToString(), 
                     				READER_COMP["DescripSimple"].ToString() 
                     			);

                            while (READER_NEW.Read())
                            {
                                CEVList.Add(READER_NEW["CEVIC"].ToString());

                            	String dsTrans =  "", dsGear =  "", dsCyl =  "", dsPts =  "", dsPass =  "", dsBrakes =  "", 
                                       dsVest =  "", dsSound =  "", dsEquip =  "", dsAir =  "", dsAirBag =  "", dsQC =  "";
                            	String nTSM = " " + READER_NEW["Descripcion"].ToString().Trim() + " ";

                                // Getting Info. from Models
                            	for (int i = 0; i < AcroDB[0].Count; i++) {
                                    if (nTSM.Contains(" " + AcroDB[0].ElementAt(i) + " ") && AcroDB[0].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[0].ElementAt(i) + " ", " ");
                            			dsTrans += AcroDB[0].ElementAt(i) + " ";
                                        //MessageBox.Show("Contains Trans: " + dsTrans);
                            		}
                                }
                                for (int i = 0; i < AcroDB[1].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[1].ElementAt(i) + " ") && AcroDB[1].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[1].ElementAt(i) + " ", " ");
                            			dsGear += AcroDB[1].ElementAt(i) + " ";
                            		}
                                }
                                for (int i = 0; i < AcroDB[2].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[2].ElementAt(i) + " ") && AcroDB[2].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[2].ElementAt(i) + " ", " ");
                            			dsCyl += AcroDB[2].ElementAt(i) + " ";
                            		}
                                }
                                for (int i = 0; i < AcroDB[3].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[3].ElementAt(i) + " ") && AcroDB[3].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[3].ElementAt(i) + " ", " ");
                            			dsPass += AcroDB[3].ElementAt(i) + " ";
                            		}
                                }
                                for (int i = 0; i < AcroDB[4].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[4].ElementAt(i) + " ") && AcroDB[4].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[4].ElementAt(i) + " ", " ");
                            			dsPts += AcroDB[4].ElementAt(i) + " ";
                            		}
                                }
                                for (int i = 0; i < AcroDB[5].Count; i++)
                                {
                            		if (nTSM.Contains(" " + AcroDB[5].ElementAt(i) + " ")  && AcroDB[5].ElementAt(i) != "NA") {
                                        nTSM = nTSM.Replace(" " + AcroDB[5].ElementAt(i) + " ", " ");
                            			dsBrakes += AcroDB[5].ElementAt(i) + " ";
                            		}
                                }
                                for (int i = 0; i < AcroDB[6].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[6].ElementAt(i) + " ") && AcroDB[6].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[6].ElementAt(i) + " ", " ");
                            			dsVest += AcroDB[6].ElementAt(i) + " ";
                            		}
                                }
                                for (int i = 0; i < AcroDB[7].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[7].ElementAt(i) + " ") && AcroDB[7].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[7].ElementAt(i) + " ", " ");
                            			dsSound += AcroDB[7].ElementAt(i) + " ";
                            		}
                                }
                                for (int i = 0; i < AcroDB[8].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[8].ElementAt(i) + " ") && AcroDB[8].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[8].ElementAt(i) + " ", " ");
                            			dsEquip += AcroDB[8].ElementAt(i) + " ";
                            		}
                                }
                                for (int i = 0; i < AcroDB[9].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[9].ElementAt(i) + " ") && AcroDB[9].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[9].ElementAt(i) + " ", " ");
                            			dsAir += AcroDB[9].ElementAt(i) + " ";
                            		}
                                }
                                for (int i = 0; i < AcroDB[10].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[10].ElementAt(i) + " ") && AcroDB[10].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[10].ElementAt(i) + " ", " ");
                            			dsAirBag += AcroDB[10].ElementAt(i) + " ";
                            		}
                                }
                                for (int i = 0; i < AcroDB[11].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[11].ElementAt(i) + " ") && AcroDB[11].ElementAt(i) != "NA")
                                    {
                            			nTSM = nTSM.Replace(" " + AcroDB[11].ElementAt(i) + " ", " ");
                            			dsQC += AcroDB[11].ElementAt(i) + " ";
                            		}
                                    
                            	}
                                
                                //MessageBox.Show("Descrip: _" + nTSM.Trim() + "_");
                                dTStand.Rows.Add(
                                	dsTrans.Trim(),
                                	dsGear.Trim(),
                                	dsCyl.Trim(),
                                	dsPts.Trim(),
                                	dsPass.Trim(),
                                	dsBrakes.Trim(),
                                	dsVest.Trim(),
                                	dsSound.Trim(),
                                	dsEquip.Trim(),
                                	dsAir.Trim(),
                                	dsAirBag.Trim(),
                                	dsQC.Trim(),
                                	nTSM.Trim()
                     			);
                            }

                            // EVALUATE SIMILARTY
                            int standRef = evalReader(dTStand);
                            // Si es mayor a 0, insertar Referencia (Modelo Equivalente)
                            if (standRef > 0)
                            {
                                myQuery = "UPDATE DatosEstandarizados SET Cia_" + numCia.ToString() + " = '" + READER_COMP["Clave"].ToString() +
                                "' WHERE CEVIC = '" + CEVList.ElementAt(standRef-1) + "' AND Modelo = '" + READER_COMP["Modelo"].ToString() + "'";
                                doQuery(myQuery
                                    ,
                                    MyConnString
                                );

                                myQuery = "";
                                //MessageBox.Show("MATCH EXITOSO EN REGISTRO", "INFO");
                                CEVList.Clear();
                            } 
                            // Sin Equivalencia, Nuevo Registro
                            else {
                                // CALCULAR CEVIC
                                Mar = (READER_COMP["Marca"].ToString().Length > 3) ? (READER_COMP["Marca"].ToString()).Substring(0, 3) : (READER_COMP["Marca"].ToString());
                                Typ = (READER_COMP["Tipo"].ToString().Length > 2) ? (READER_COMP["Tipo"].ToString()).Substring(0, 2) : (READER_COMP["Tipo"].ToString());
                                Mod = READER_COMP["Modelo"].ToString();
                                cveCo = READER_COMP["Clave"].ToString();
                                cveCEVIC = Mar + Typ + Mod + cveCo + "_X";

                                // COMPROBAR SI CEVIC ESTÁ EN LA TABLA
                                try
                                {
                                    OleDbConnection CONNECT3 = new OleDbConnection(MyConnString);
                                    CONNECT3.Open();
                                    // SELECT COUNT (*) FROM DatosEstandarizados WHERE CEVIC LIKE '8514_X??'
                                    OleDbCommand COMMAND_CEVIC = new OleDbCommand("SELECT COUNT (*) FROM DatosEstandarizados WHERE CEVIC LIKE '" + cveCEVIC + "__'", CONNECT);
                                    nMod = (Int32)COMMAND_CEVIC.ExecuteScalar();
                                    CONNECT3.Close();

                                    cveCEVIC += nMod.ToString("D2");
                                    myQuery = "INSERT INTO DatosEstandarizados" +
                                        // Cambiar Nº de Compañia
                                            "(Cia_" + numCia.ToString() + ", CEVIC, Modelo, CveMarca_Cia, CveTipo_Cia, CveVersion_Cia, CveTrans_Cia, Marca, Tipo, Descripcion)" +
                                            "VALUES ('" +
                                            cveCo + "','" +
                                            cveCEVIC + "','" +
                                            Mod +
                                            "','','','','','" +
                                            READER_COMP["Marca"].ToString() + "','" +
                                            READER_COMP["Tipo"].ToString() + "','" +
                                            READER_COMP["DescripTSM"].ToString() + "')";

                                    doQuery(myQuery
                                    ,
                                    MyConnString
                                    );

                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.ToString());
                                }
                                myQuery = "";
                                cveCEVIC = "";

                                //break;
                                //MessageBox.Show("New");
                            }
                        }
                        else
                        {
                            Mar = (READER_COMP["Marca"].ToString().Length > 3) ? (READER_COMP["Marca"].ToString()).Substring(0, 3) : (READER_COMP["Marca"].ToString());
                            Typ = (READER_COMP["Tipo"].ToString().Length > 2) ? (READER_COMP["Tipo"].ToString()).Substring(0, 2) : (READER_COMP["Tipo"].ToString());
                            Mod = READER_COMP["Modelo"].ToString();
                            nMod = 0;
                            cveCo = READER_COMP["Clave"].ToString();
                            cveCEVIC = Mar + Typ + Mod + cveCo + "_X" + nMod.ToString("D2");

                            myQuery = "INSERT INTO DatosEstandarizados" +
                                // Cambiar Nº de Compañia
                                "(Cia_" + numCia.ToString() + ", CEVIC, Modelo, CveMarca_Cia, CveTipo_Cia, CveVersion_Cia, CveTrans_Cia, Marca, Tipo, Descripcion)" +
                                "VALUES ('" +

                                cveCo + "','" +
                                cveCEVIC + "','" +
                                Mod +
                                "','','','','','" +
                                READER_COMP["Marca"].ToString() + "','" +
                                READER_COMP["Tipo"].ToString() + "','" +
                                READER_COMP["DescripTSM"].ToString() + "')";

                            doQuery(myQuery
                                ,
                                MyConnString
                            );

                            myQuery = "";
                            cveCEVIC = "";
                            //MessageBox.Show("NUEVO REGISTRO, NUEVO AUTO", "INFO");
                        }
                        // Progress Update
                        pBCount++;
                        this.Invoke(this.updateStatusDelegate);
                        
                        READER_NEW.Close();
                        CONNECT2.Close();
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.ToString(), "ERROR ON INSERT ON DATOS ESTANDARIZADOS");
                    }
                }
                READER_COMP.Close();
                CONNECT.Close();
                MessageBox.Show("SUCCESS");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString(), "ERROR ON GETTING FROM COMPANY");
            }
            finally {
                this.workerThread.Abort();
            }
        }

        // METODOS PARA LA REVISION DE HOMOLOGACIÓN
        // SUSTITUIR STRING EN DESCRIPCION
        private void sustitString(String sOld, String sNew, String sDB, String sMarca, String sType, String sField)
        {
            OleDbConnection CONNECT = new OleDbConnection(
                MyConnString);
            String CONSULT;

            if (sType == "false")          
                CONSULT = "SELECT * FROM D_" + sDB + " WHERE Marca = '" + sMarca + "'";
            else
                CONSULT = "SELECT * FROM D_" + sDB + " WHERE Marca = '" + sMarca + "' AND Tipo = '" + sType + "'";
            
            try {
                CONNECT.Open();
                OleDbCommand COMMAND = new OleDbCommand(CONSULT, CONNECT);
                OleDbDataReader READER = COMMAND.ExecuteReader();
                int i = 0;

                if (READER.HasRows) { 
                    while(READER.Read()) {
                        String x = " " + READER[sField].ToString() + " ";

                        if (x.Contains(sOld)) {
                            //MessageBox.Show(READER["Clave"].ToString() + ": " + x, "oLD");
                            x = x.Replace(sOld, " " + sNew + " ");

                            if (sType == "false")
                            {
                                doQuery("UPDATE D_" + sDB + " SET " + sField + " = '" + x.Trim() + "' WHERE Clave = '" + READER["Clave"].ToString() +
                                    "' AND Marca = '" + READER["Marca"].ToString() + "'",

                                    MyConnString
                                    );
                            }
                            else {
                                doQuery("UPDATE D_" + sDB + " SET " + sField + " = '" + x.Trim() + "' WHERE Clave = '" + READER["Clave"].ToString() +
                                   "' AND Marca = '" + READER["Marca"].ToString() + "' AND Tipo = '" + READER["Tipo"].ToString() + "'",

                                   MyConnString
                                   );
                            
                            
                            }
                            //MessageBox.Show(READER["Clave"].ToString() + ": " + x, "nEW");
                            i++;
                        }
                        
                    }
                }
                MessageBox.Show("SUCCESS: " + i.ToString(), "INFO");
                READER.Close();
                CONNECT.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString(), "ERROR");
            }
        }

        // BORRAR STRING EN DESCRIPCION
        private void deletString(String sOld, String sDB, String sMarca, String sField)
        {
            OleDbConnection CONNECT = new OleDbConnection(
                MyConnString);
            String CONSULT = "SELECT * FROM D_" + sDB + " WHERE Marca = '" + sMarca + "'";
            try
            {
                CONNECT.Open();
                OleDbCommand COMMAND = new OleDbCommand(CONSULT, CONNECT);
                OleDbDataReader READER = COMMAND.ExecuteReader();
                int i = 0;
                if (READER.HasRows)
                {
                    while (READER.Read())
                    {
                        String x = " " + READER[sField].ToString().Trim() + " ";
                        if (x.Contains(sOld))
                        {
                            //MessageBox.Show(READER["Clave"].ToString() + ": " + x, "oLD");
                            x = x.Replace(sOld, " ");
                            doQuery("UPDATE D_" + sDB + " SET " + sField + " = '" + x.Trim() + "' WHERE Clave = '" + READER["Clave"].ToString() +
                                "' AND Marca = '" + READER["Marca"].ToString() + "'",

                                MyConnString
                                );
                            //MessageBox.Show(READER["Clave"].ToString() + ": " + x, "nEW");
                            i++;
                        }

                    }
                }
                MessageBox.Show("SUCCESS: " + i.ToString(), "INFO");
                READER.Close();
                CONNECT.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString(), "ERROR");
            }
        }

        // INSERTAR EN CAMPO DESDE DESCRIPCION 
        private void insertStringField(String sOld, String sNew, String sField, String sDB, String sMarca)
        {
            OleDbConnection CONNECT = new OleDbConnection(
                MyConnString);
            String CONSULT = "SELECT * FROM D_" + sDB + " WHERE Marca = '" + sMarca + "'";
            try
            {
                CONNECT.Open();
                OleDbCommand COMMAND = new OleDbCommand(CONSULT, CONNECT);
                OleDbDataReader READER = COMMAND.ExecuteReader();
                int i = 0;
                if (READER.HasRows)
                {
                    while (READER.Read())
                    {
                        String x = " " + READER["DescripSimple"].ToString().Trim() + " ";
                        String y = READER[sField].ToString();
                        if (x.Contains(sOld))
                        {
                            //MessageBox.Show(READER["Clave"].ToString() + ": " + x, "oLD");
                            x = x.Replace(sOld, " ");
                            y = (y.Length > 0) ? y + " " + sNew : sNew;

                            doQuery("UPDATE D_" + sDB + " SET DescripSimple = '" + x.Trim() + "', "+ sField +" = '" +  y + "' WHERE Clave = '" + READER["Clave"].ToString() +
                                "' AND Marca = '" + READER["Marca"].ToString() + "'",

                                MyConnString
                                );
                            i++;
                        }

                    }
                }
                MessageBox.Show("SUCCESS: " + i.ToString(), "INFO");
                READER.Close();
                CONNECT.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString(), "ERROR");
            }
        }



        // METODOS PARA HOMOLOGAR
        private void subsAcronym(String sOld, String sNew, String sField, String sDB)
        {
            OleDbConnection CONNECT = new OleDbConnection(
                MyConnString);
            String CONSULT = "SELECT * FROM D_" + sDB + " ORDER BY Clave, Modelo";
            try
            {
                CONNECT.Open();
                OleDbCommand COMMAND = new OleDbCommand(CONSULT, CONNECT);
                OleDbDataReader READER = COMMAND.ExecuteReader();
                int i = 0;

                while (READER.Read())
                {
                    String x = " " + READER["DescripSimple"].ToString() + " ";
                    String y = READER[sField].ToString();

                    if (x.Contains(sOld))
                    {
                        if (!(" " + y.Trim() + " ").Contains(" " + sNew + " "))
                        {
                            //MessageBox.Show(READER["Clave"].ToString() + ": " + x, "oLD");
                            x = x.Replace(sOld, " ");
                            y = (y.Length > 0) ? y + " " + sNew: sNew;
                            doQuery("UPDATE D_" + sDB + " SET DescripSimple = '" + x.Trim() + "', "+ sField +" = '" + y + "' WHERE Clave = '" + READER["Clave"].ToString() +
                                "' AND Modelo = " + READER["Modelo"].ToString(),

                                MyConnString
                                 );
                            //MessageBox.Show(READER["Clave"].ToString() + ": " + x, "nEW");
                            
                        } else {
                            x = x.Replace(sOld, " ");
                            doQuery("UPDATE D_" + sDB + " SET DescripSimple = '" + x.Trim() + "' WHERE Clave = '" + READER["Clave"].ToString() +
                                "' AND Modelo = " + READER["Modelo"].ToString(),

                                MyConnString
                                 );
                        }
                        i++;
                    }

                }
                MessageBox.Show("SUCCESS: " + i.ToString(), "INFO");
                READER.Close();
                CONNECT.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString(), "ERROR");
            }
        }

        private void subsAcronymSimple(String sOld, String sNew, String sDB)
        {
            OleDbConnection CONNECT = new OleDbConnection(
                MyConnString);
            String CONSULT = "SELECT * FROM D_" + sDB;
            try
            {
                CONNECT.Open();
                OleDbCommand COMMAND = new OleDbCommand(CONSULT, CONNECT);
                OleDbDataReader READER = COMMAND.ExecuteReader();
                int i = 0;

                while (READER.Read())
                {
                    String x = " " + READER["DescripSimple"].ToString() + " ";

                    if (x.Contains(sOld))
                    {
                        //MessageBox.Show(READER["Clave"].ToString() + ": " + x, "oLD");
                        x = x.Replace(sOld, sNew);
                        doQuery("UPDATE D_" + sDB + " SET DescripSimple = '" + x.Trim() + "' WHERE Clave = '" + READER["Clave"].ToString() +
                            "' AND Modelo = " + READER["Modelo"].ToString(),

                            MyConnString
                             );
                        //MessageBox.Show(READER["Clave"].ToString() + ": " + x, "nEW");
                        i++;
                    }

                }
                MessageBox.Show("SUCCESS: " + i.ToString(), "INFO");
                READER.Close();
                CONNECT.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString(), "ERROR");
            }
        }

        private void changeType(String tOld, String tNew, String tBrand, String sDB, Boolean allType)
        {
            OleDbConnection CONNECT = new OleDbConnection(
                MyConnString);
            String CONSULT = "SELECT * FROM D_" + sDB + " WHERE Marca = '" + tBrand + "'";
            try
            {
                CONNECT.Open();
                OleDbCommand COMMAND;
                if (allType)
                    COMMAND = new OleDbCommand(CONSULT, CONNECT);
                else
                    COMMAND = new OleDbCommand(CONSULT + " AND Tipo = '" + tOld.Trim() + "'", CONNECT);

                OleDbDataReader READER = COMMAND.ExecuteReader();
                int i = 0;

                while (READER.Read())
                {
                    String x = " " + READER["DescripSimple"].ToString().Trim() + " ";

                    if (allType)
                    {
                        if (x.Contains(tOld))
                        {
                            //REMOVER AL INICIO
                            
                            //int n = x.IndexOf(tOld);
                            //x = x.Remove(n, tOld.Length);
                            x = x.Replace(tOld, " ");

                            doQuery("UPDATE D_" + sDB + " SET Tipo = '" + tNew + "', DescripSimple = '" + x.Trim() + "' WHERE Clave = '" + READER["Clave"].ToString() +
                                "' AND Marca = '" + READER["Marca"].ToString() + "'",

                                MyConnString
                                 );
                            i++;
                        }
                    }
                    else
                    {
                        doQuery("UPDATE D_" + sDB + " SET Tipo = '" + tNew + "' WHERE Tipo = '" + tOld.Trim() +
                                "' AND Marca = '" + READER["Marca"].ToString() + "' AND Tipo <> '" + tNew + "'",

                                MyConnString
                                 );
                        i++;
                    }

                }
                MessageBox.Show("SUCCESS: " + i.ToString(), "INFO");
                READER.Close();
                CONNECT.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString(), "ERROR");
            }
        }

        private void deletStringAll(String sOld, String sDB, String sField)
        {
            OleDbConnection CONNECT = new OleDbConnection(
                MyConnString);
            String CONSULT = "SELECT * FROM D_" + sDB;
            try
            {
                CONNECT.Open();
                OleDbCommand COMMAND = new OleDbCommand(CONSULT, CONNECT);
                OleDbDataReader READER = COMMAND.ExecuteReader();
                int i = 0;
                if (READER.HasRows)
                {
                    while (READER.Read())
                    {
                        String x = " " + READER[sField].ToString().Trim() + " ";
                        if (x.Contains(sOld))
                        {
                            //MessageBox.Show(READER["Clave"].ToString() + ": " + x, "oLD");
                            x = x.Replace(sOld, " ");
                            doQuery("UPDATE D_" + sDB + " SET " + sField + " = '" + x.Trim() + "' WHERE Clave = '" + READER["Clave"].ToString() +
                                "' AND Modelo = " + READER["Modelo"].ToString(),

                                MyConnString
                                );
                            //MessageBox.Show(READER["Clave"].ToString() + ": " + x, "nEW");
                            i++;
                        }

                    }
                }
                MessageBox.Show("SUCCESS: " + i.ToString(), "INFO");
                READER.Close();
                CONNECT.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString(), "ERROR");
            }
        }


        // METODOS ABA
       
        // FUNCION PARA ABA
        private void setDoors(String sDB)
        {
            OleDbConnection CONNECT = new OleDbConnection(
                MyConnString);
            String CONSULT = "SELECT * FROM D_" + sDB;
            try
            {
                CONNECT.Open();
                OleDbCommand COMMAND;
                COMMAND = new OleDbCommand(CONSULT, CONNECT);

                OleDbDataReader READER = COMMAND.ExecuteReader();
                int i = 0;

                while (READER.Read())
                {
                    String x = READER["DescripSimple"].ToString().Trim();
                    String y = (x.Length > 1) ? x.Substring(x.Length - 1, 1): x;
                    
                    int numP = 0;
                    if (x.Contains(y) && Int32.TryParse(y, out numP))
                    {
                 
                        if (6 > numP && numP > 0)
                        {
                            x = (x.Length > 1) ? x.Remove(x.Length - 1, 1): "";
                            doQuery("UPDATE D_" + sDB + " SET DescripSimple = '" + x.Trim() + "', Puertas = '" + numP.ToString() + "Ptas' WHERE Clave = '" + READER["Clave"].ToString() +
                                "' AND Modelo = " + READER["Modelo"].ToString() + "",

                                MyConnString
                                 );
                            i++;
                        }

                    }
                }
                MessageBox.Show("SUCCESS: " + i.ToString(), "INFO");
                READER.Close();
                CONNECT.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString(), "ERROR");
            }
        }

        private void removeBrandType(String sDB)
        {
            OleDbConnection CONNECT = new OleDbConnection(
                MyConnString);
            String CONSULT = "SELECT * FROM D_" + sDB;
            try
            {
                CONNECT.Open();
                OleDbCommand COMMAND = new OleDbCommand(CONSULT, CONNECT);
                OleDbDataReader READER = COMMAND.ExecuteReader();
                int i = 0;

                while (READER.Read())
                {
                    String x = READER["DescripSimple"].ToString().TrimStart();
                    //String y = READER["Marca"].ToString();
                    String z = READER["Tipo"].ToString();
                    if (x.Contains(z))
                    {
                        //REMOVER AL INICIO
                        x = x.Remove(0, z.Length);
                        doQuery("UPDATE D_" + sDB + " SET DescripSimple = '" + x.Trim() + "' WHERE Clave = '" + READER["Clave"].ToString() +
                            "' AND Marca = '" + READER["Marca"].ToString() + "' AND Tipo = '" + READER["Tipo"].ToString() + "'",

                            MyConnString
                             );
                        i++;
                    }

                }
                MessageBox.Show("SUCCESS: " + i.ToString(), "INFO");
                READER.Close();
                CONNECT.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString(), "ERROR");
            }
        }

        // INTERFAZ REVISION 
        private void btnDelete1_Click_1(object sender, EventArgs e)
        {
            if (txtOldString1.Text != "" && txtNewString1.Text == "")
            {
                DialogResult dialogResult = MessageBox.Show("¿Borrar contenido especificado?", "Confirmación", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    if (!chkFieldD.Checked && cmbField1.SelectedIndex == -1)
                    {
                        deletString(" " + txtOldString1.Text + " ", cmbCompany1.SelectedItem.ToString(), cmbBrand1.SelectedItem.ToString().ToUpper(), "DescripSimple");
                    }
                    else
                    {
                        if (cmbField1.SelectedIndex <= 0)
                        {
                            deletString(" " + txtOldString1.Text + " ", cmbCompany1.SelectedItem.ToString(), cmbBrand1.SelectedItem.ToString().ToUpper(), cmbField1.SelectedItem.ToString());
                            chkFieldD.Checked = !chkFieldD.Checked;
                        }
                    } 
                }
            }
            else
            {
                MessageBox.Show("You eat it at deleting", "ERROR");
            }
            cmbField1.SelectedIndex = -1;
        }

        private void btnSubDescrip_Click_1(object sender, EventArgs e)
        {
            if (txtNewString1.Text != "" && cmbBrand1.SelectedIndex >= 0)
            {
                if (!chkFieldD.Checked && cmbField1.SelectedIndex == -1)
                {
                    sustitString(" " + txtOldString1.Text +  " ", txtNewString1.Text, cmbCompany1.SelectedItem.ToString(), cmbBrand1.SelectedItem.ToString(), "false", "DescripSimple");
                    
                }
                else
                {
                    if (cmbField1.SelectedIndex <= 0)
                    {
                        sustitString(" " + txtOldString1.Text + " ", txtNewString1.Text, cmbCompany1.SelectedItem.ToString(), cmbBrand1.SelectedItem.ToString().ToUpper(), "false", cmbField1.SelectedItem.ToString());
                        chkFieldD.Checked = !chkFieldD.Checked;                     
                    }
                }
                cmbType1.SelectedIndex = -1;

            }
            else
            {
                MessageBox.Show("You eat it at substituting an string, BABOSO", "ERROR");
            }
            txtNewString1.Text = "";
            cmbField1.SelectedIndex = -1;
            cmbType1.SelectedIndex = -1;
        }

        private void btnDescripToField_Click_1(object sender, EventArgs e)
        {
            // insertStringField(String sOld, String sNew, String sField, String sDB, String sMarca)
            if (txtOldString1.Text != "" && cmbField1.SelectedIndex >= 0 && chkFieldD.Checked)
            {
                insertStringField(" " + txtOldString1.Text + " ", txtNewString1.Text, cmbField1.SelectedItem.ToString(), cmbCompany1.SelectedItem.ToString(), cmbBrand1.SelectedItem.ToString().ToUpper());
            }
            else
            {
                MessageBox.Show("You eat it at inserting on new field", "ERROR");
            }
            txtNewString1.Text = "";
            cmbField1.SelectedIndex = -1;
            chkFieldD.Checked = false;
        }

        // HOMOLOG
        // CONTENIDO DE HOMOLOGACION
        private void cmbField0_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Trans, Cil, Vest, Aire, QC, Equipado, EE, BAire, Sonido, ABS, RA, FN, CodRaro, DH 
            String[,] fieldOptions = { 
                                        { "2", "3", "4", "5", "6", "7", "8", "9", "10", "12", 
                                           "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", 
                                           "23", "24", "25", "27", "30", 
                                           "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
                                        { "2Ptas", "2y4Ptas", "3Ptas", "3y5Ptas", "4Ptas", "5Ptas", "", "", "", "", "", "", "", "","", "", "", "","", "", "", "", "",
                                            "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
                                        { "Aut", "Std", "7G-DCT", "7G-Tronic", "9G-DCT", "9G-Tronic", "AMG-Speedshift", "ASG", "AutoStick", "CVT", "DCT", 
                                           "Drivelogic", "DSG", "Dualogic", "DuoSelect", "Easytronic", "EDC", "Geartronic", "GETRAG", "G-Tronic", "HSD", 
                                           "Hydramatic", "Lineartronic", "M-DKG", "MCT", "Multitronic", "PDK", "Powershift", "Q-Tronic", "R-Tronic", "Secuencial", "SelectShift",
                                           "Selespeed", "Sentronic", "Shiftmatic", "Shiftronic", "SMG-II", "SMG", "Sportronic", "SportShift", "Steptronic", "S-Tronic",
                                           "TCT", "Tiptronic", "Touchtronic", "X-Tronic"}, 
                                        { "0Cil", "2Cil", "3Cil", "4Cil", "5Cil", "6Cil", "8Cil", "10Cil", "12Cil", 
                                            "", "", "", "", "", "", "", "","", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
                                        { "Alcantara", "Gamusina", "Gamuza", "Leatherette", "Napa", "Piel parcial", "Piel", "Tela", "Terciopelo", "Velour", 
                                            "Vinil", "", "", "","", "", "", "","", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
                                        { "C/AAcc", "S/AAcc", "", "", "", "", "", "", "", "", "", "", "", "","", "", "", "","", "", "", "", "",
                                            "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
                                        { "C/Qcc", "S/Qcc", "", "", "", "", "", "", "", "", "", "", "", "","", "", "", "","", "", "", "", "",
                                            "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
                                        { "6CD", "AFS", "ASIST.EST.", "AUDIO MANEJO", "C/CAMARA", "C/LOCKER", "CAM.TRAS.", "COMAND ONLINE", "COMP.VIAJE.", "CTROL/AUDIO", 
                                            "CTROL/VOZ", "EQUIPADO", "F/BI-XENON", "F/XENON", "GETRONIC", "GMLINK", "FULL LINK", "GPS", "HIELERA", "HILL HOLDER", "JOYBOX", 
                                            "MEDIA NAV", "MULTIMEDIA", "MYGIG", "NAVIGON", "PTA.TRAS.ELEC.", "SEMIEQUIPADO", "RNS-510", "SIST.ENTRET.", "SIST.NAV.", 
                                            "TPM", "TV", "UCONNECT", "WIFI", "", "", "", "", "", "", "", "", "", "", "", ""},
                                        { "CE", "EE", "SE", "", "", "", "", "", "", "", "", "", "", "","", "", "", "","", "", "", "", "",
                                            "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
                                        { "CB", "CBL", "SB", "", "", "", "", "", "", "", "", "", "", "","", "", "", "","", "", "", "", "",
                                            "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
                                        { "AM", "BD", "BOSE", "BT", "CD", "CT", "DVD", "DYNAUDIO", "FENDER", "FM",
                                            "MP3", "RADIO", "SS", "USB", "","", "", "","", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
                                        { "ABS", "D/ABS", "D/V", "D/T", "DIS", "NEU", "TAM", "V", "V/DIS", "V/T", "", "", "", "","", "", "", "","", "", "", "", "",
                                            "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
                                        { "R-13", "R-14", "R-15", "R-16", "R-17", "R-18", "R-19", "R-20", "R-21", "R-22",
                                            "R-25", "RA", "RA-14", "RA-15", "RA-16", "RA-17", "RA-18", "RA-19", "RA-20",
                                            "RA-21", "", "", "", "", "", "", "", "", "", 
                                            "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
                                        { "FN", "", "", "", "", "", "", "", "", "", "", "", "", "","", "", "", "","", "", "", "", "",
                                            "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
                                        { "SM", "", "", "", "", "", "", "", "", "", "", "", "", "","", "", "", "","", "", "", "", "",
                                            "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
                                        { "DH", "DHS", "", "", "", "", "", "", "", "", "", "", "", "","", "", "", "","", "", "", "", "",
                                            "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""}
                                     };


            cmbDescrip0.Items.Clear();
            cmbDescrip0.Text = "";
            if (cmbField0.SelectedIndex >= 0)
            {
                for (int i = 0; i < 45; i++)
                {
                    if (fieldOptions[cmbField0.SelectedIndex, i] != "")
                        cmbDescrip0.Items.Add(fieldOptions[cmbField0.SelectedIndex, i]);
                }
            }
        }

        private void btnDescrip0_Click(object sender, EventArgs e)
        {
            if (cmbCompany0.SelectedIndex >= 0)
            {
                if (txtOld0.Text != "")
                {
                    if (chkSimpleAcr.Checked && txtNew0.Text != "")
                    {
                        subsAcronymSimple(" " + txtOld0.Text.Trim() + " ", " " + txtNew0.Text.Trim() + " ", cmbCompany0.SelectedItem.ToString());
                        chkSimpleAcr.Checked = !chkSimpleAcr.Checked;
                        txtOld0.Clear();
                        txtNew0.Clear();
                    }
                    else
                    {
                        if (cmbField0.SelectedIndex >= 0 && cmbDescrip0.SelectedIndex >= 0)
                        {
                            subsAcronym(" " + txtOld0.Text + " ", cmbDescrip0.SelectedItem.ToString(), cmbField0.SelectedItem.ToString(), cmbCompany0.SelectedItem.ToString());
                            txtOld0.Clear();

                        }
                        else
                        {
                            MessageBox.Show("You eat it, baboso!", "ERROR");
                        }
                    }
                }
                else {
                    MessageBox.Show("Escribe String a sustituir", "ERROR");
                }
            }
            else {
                MessageBox.Show("Selecciona una compañía", "ERROR");
            }
        }

        private void btnType0_Click(object sender, EventArgs e)
        {
            //int INDEX = cmbBrand0.SelectedIndex;
            if (cmbBrand0.SelectedIndex >= 0 && txtOld0.Text != "" && txtNew0.Text != "")
            {
                if (chkVoidType.Checked)
                {
                    changeType(" " + txtOld0.Text + " ", txtNew0.Text, cmbBrand0.SelectedItem.ToString(), cmbCompany0.SelectedItem.ToString(), true);
                    txtNew0.Clear();
                }
                else {
                    changeType(" " + txtOld0.Text + " ", txtNew0.Text, cmbBrand0.SelectedItem.ToString(), cmbCompany0.SelectedItem.ToString(), false);
                    txtNew0.Clear();
                } /*
                if (cmbCompany0.SelectedItem == cmbCompany1.SelectedItem)
                {
                    cmbBrand1.SelectedIndex = -1;
                    cmbBrand1.SelectedIndex = INDEX;
                } */
            }
            else {
                MessageBox.Show("Selecciona la marca y escribe el tipo", "ERROR");
            }
        }

        private void chkFieldD_CheckedChanged(object sender, EventArgs e)
        {
            if (!chkFieldD.Checked)
                cmbField1.SelectedIndex = -1;
        }

        private void cmbBrand1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbCompany1.SelectedIndex >= 0)
            {
                cmbType1.Items.Clear();
                OleDbConnection CONNECT = new OleDbConnection(
                    MyConnString);
                String CONSULT = "SELECT DISTINCT Tipo as Type FROM D_" + cmbCompany1.SelectedItem.ToString() + " WHERE Marca = '" + cmbBrand1.SelectedItem.ToString() + "'";
                try
                {
                    CONNECT.Open();
                    OleDbCommand COMMAND = new OleDbCommand(CONSULT, CONNECT);
                    OleDbDataReader READER = COMMAND.ExecuteReader();

                    if (READER.HasRows)
                    {
                        while (READER.Read())
                        {
                            if (READER["Type"].ToString() != "")
                                cmbType1.Items.Add(READER["Type"]);
                        }
                    }
                    // MessageBox.Show("Type load successful ", "INFO");
                    READER.Close();
                    CONNECT.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), "ERROR");
                }
            }
            else {
                MessageBox.Show("Selecciona una Compañía", "ERROR");
                cmbBrand1.SelectedIndex = -1;
            }
        }

        private void cmbCompany1_SelectedIndexChanged(object sender, EventArgs e)
        {
            cmbBrand1.Items.Clear();
            String brand = cmbBrand1.Text;
            cmbBrand1.SelectedIndex = -1;
            cmbBrand1.Text = "";
            OleDbConnection CONNECT = new OleDbConnection(
                MyConnString);
            String CONSULT = "SELECT DISTINCT Marca as Brand FROM D_" + cmbCompany1.SelectedItem.ToString();
            try
            {
                CONNECT.Open();
                OleDbCommand COMMAND = new OleDbCommand(CONSULT, CONNECT);
                OleDbDataReader READER = COMMAND.ExecuteReader();

                if (READER.HasRows)
                {
                    while (READER.Read())
                    {
                        if (READER["Brand"].ToString() != "")
                            cmbBrand1.Items.Add(READER["Brand"]);
                    }
                }
                // MessageBox.Show("Type load successful ", "INFO");
                READER.Close();
                CONNECT.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "ERROR");
            }
            finally
            {
                if (cmbBrand1.Items.Contains(brand))
                {
                    cmbBrand1.SelectedItem = brand;
                }
            }
        }

        private void cmbCompany0_SelectedIndexChanged(object sender, EventArgs e)
        {
            cmbBrand0.Items.Clear();
            String brand = cmbBrand0.Text;
            cmbBrand0.SelectedIndex = -1;
            cmbBrand0.Text = "";
            OleDbConnection CONNECT = new OleDbConnection(
                MyConnString);
            String CONSULT = "SELECT DISTINCT Marca as Brand FROM D_" + cmbCompany0.SelectedItem.ToString();
            try
            {
                CONNECT.Open();
                OleDbCommand COMMAND = new OleDbCommand(CONSULT, CONNECT);
                OleDbDataReader READER = COMMAND.ExecuteReader();

                if (READER.HasRows)
                {
                    while (READER.Read())
                    {
                        if (READER["Brand"].ToString() != "")
                            cmbBrand0.Items.Add(READER["Brand"]);
                    }
                }
                // MessageBox.Show("Type load successful ", "INFO");
                READER.Close();
                CONNECT.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "ERROR");
            }
            finally {
                if (cmbBrand0.Items.Contains(brand)) {
                    cmbBrand0.SelectedItem = brand;
                }
            }
        }

        private void cmbBrand0_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbCompany0.SelectedIndex < 0) {
                cmbBrand0.Items.Clear();
                MessageBox.Show("Selecciona una compañia", "ERROR");
                cmbBrand0.SelectedIndex = -1;
            }
        }

        private void btnSubDescripType_Click(object sender, EventArgs e)
        {
            if (txtNewString1.Text != "" && cmbType1.SelectedIndex >= 0)
            {
                if (!chkFieldD.Checked && cmbField0.SelectedIndex == -1)
                {
                    sustitString(" " + txtOldString1.Text + " ", txtNewString1.Text, cmbCompany1.SelectedItem.ToString(), cmbBrand1.SelectedItem.ToString(), cmbType1.SelectedItem.ToString(), "DescripSimple");

                }
                else
                {
                    if (cmbField0.SelectedIndex <= 0)
                    {
                        sustitString(" " + txtOldString1.Text + " ", txtNewString1.Text, cmbCompany1.SelectedItem.ToString(), cmbBrand1.SelectedItem.ToString().ToUpper(), cmbType1.SelectedItem.ToString(), cmbField1.SelectedItem.ToString());
                        chkFieldD.Checked = !chkFieldD.Checked;
                    }
                }

            }
            else
            {
                MessageBox.Show("Selecciona Tipo y escribe un string, baboso!", "ERROR");
            }
            txtNewString1.Text = "";
            cmbField1.SelectedIndex = -1;
            cmbType1.SelectedIndex = -1;
        }

        private void chkSimpleAcr_CheckedChanged(object sender, EventArgs e)
        {
            cmbField0.SelectedIndex = -1;
            cmbDescrip0.SelectedIndex = -1;
        }

        public Tuple<int, int> getIndex(double[,] jaggedArray, double value)
        { 
            int w = jaggedArray.GetLength(0); // width
            int h = jaggedArray.GetLength(1); // height

            for (int x = 0; x < w; ++x)
            {
                for (int y = 0; y < h; ++y)
                {
                    if (jaggedArray[x, y].Equals(value))
                        return Tuple.Create(x, y);
                }
            }

            return Tuple.Create(-1, -1);
        }

        // Playground
        private void btnTest_Click_1(object sender, EventArgs e)
        {
            if (chk_Test.Checked)
            {
                String z1 = txtTest1.Text;
                String z = txtTest2.Text;

                var jw = new JaroWinkler();
                var lv = new Levenstein();
                //List<FuzzyStringComparisonOptions> options = new List<FuzzyStringComparisonOptions>();
                //options.Add(FuzzyStringComparisonOptions.UseJaroWinklerDistance);

                //MessageBox.Show(jw.Similarity(z1, z).ToString());
                //MessageBox.Show(z1.ApproximatelyEquals(z, options, FuzzyStringComparisonTolerance.Strong).ToString());
                //MessageBox.Show("JW: " + jw.GetSimilarity(z1, z).ToString() + ", LV: " + Levenshtein.CalculateDistance(z1, z, 1).ToString() + ", LV2: " + lv.GetSimilarity(z1, z).ToString());

                double[,] Arry = { { 1.0, 2.25, 4.64 }, { 4.23, 5.0, 4.65}, {1.5, 4.7, 4.78}, {1.2, 5.6513, 5.6529} };
                double Max = Arry.Cast<double>().Max();
                Tuple<int, int> position = getIndex(Arry, Max);
                MessageBox.Show(position.ToString(), Max.ToString());

            }
            else {
                if (txtTest1.Text == "" && txtTest2.Text == "")
                {
                    // DO SOMETHING HERE
                    OleDbConnection CONNECT = new OleDbConnection(
                     "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Users\\Griselda\\Desktop\\Qualitas22022018.mdb;");
                    String CONSULTA = "SELECT * FROM Tar_Qualitas WHERE cMarca = 'MO' ORDER BY cCAMIS";
                    try
                    {
                        CONNECT.Open();
                        OleDbCommand COMMAND = new OleDbCommand(CONSULTA, CONNECT);
                        OleDbDataReader READER = COMMAND.ExecuteReader();

                        while (READER.Read())
                        {
                            doQuery("UPDATE D_Qualitas" +
                                " SET DescripTSM = 'MO'" +
                                " WHERE Clave = '" + Convert.ToInt32(READER["cCAMIS"].ToString()).ToString() +
                                "' AND Modelo = " +
                                READER["cModelo"].ToString(),
                                MyConnString
                            );

                        }
                        READER.Close();
                        CONNECT.Close();
                        MessageBox.Show("SUCCESS");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString(), "ERROR");
                    }

                }


                
            }
        }

        // INTERFAZ PROCESOS DE HOMOLOGACIÓN
        private void cmbCompany2_SelectedIndexChanged(object sender, EventArgs e)
        {
            int[] num_Comp = {7, 21, 5, 26, 20, 2, 12};
            txtCompID2.Text = num_Comp[cmbCompany2.SelectedIndex].ToString();
            //pbarProgress.Value = 0;
        }

        private void btnAddCompany2_Click(object sender, EventArgs e)
        {
            cmbCompany2.Enabled = false;
            
            timeStart = DateTime.Now;

            if (cmbCompany2.SelectedIndex >= 0)
            {
                int nComp = Convert.ToInt32(txtCompID2.Text);
                string nmComp = cmbCompany2.SelectedItem.ToString();

                int count = 0;
                try
                {
                    OleDbConnection CONNECT_COUNT = new OleDbConnection(MyConnString);
                    CONNECT_COUNT.Open();
                    OleDbCommand COMMAND_COUNT = new OleDbCommand("SELECT COUNT (Cia_" + nComp.ToString() + ") FROM DatosEstandarizados WHERE (Cia_" + nComp.ToString() + " <> '')", CONNECT_COUNT);

                    count = Convert.ToInt32(COMMAND_COUNT.ExecuteScalar());

                    CONNECT_COUNT.Close();
                }
                catch (Exception Ex)
                {
                    MessageBox.Show(Ex.ToString() + " COMP:" +nmComp.ToString(), "Error en Count de Registros");
                }

                if (count == 0)
                {
                    this.workerThread = null;
                    this.workerThread = new Thread(
                            () => insertNewComp(nComp, nmComp)
                        );
                    this.workerThread.Start();
                }
                else {
                    MessageBox.Show("La Tabla ya esta en la Homologacion", "Error");
                }
            }
            else {
                MessageBox.Show("Selecciona una Compañía", "ERROR");
            }

            cmbCompany2.Enabled = true;
        }

        private void btnGenerateStandard_Click(object sender, EventArgs e)
        {
            cmbCompany2.Enabled = false;
            if (cmbCompany2.SelectedIndex >= 0)
            {
                try
                {
                    String Comp = cmbCompany2.SelectedItem.ToString();
                    timeStart = DateTime.Now;

                    int count = 0;
                    try
                    {
                        OleDbConnection CONNECT_COUNT = new OleDbConnection(MyConnString);
                        CONNECT_COUNT.Open();
                        OleDbCommand COMMAND_COUNT = new OleDbCommand("SELECT COUNT (*) FROM Estandarizados_" + Comp, CONNECT_COUNT);

                        count = Convert.ToInt32(COMMAND_COUNT.ExecuteScalar());

                        CONNECT_COUNT.Close();
                    }
                    catch (Exception Ex)
                    {
                        MessageBox.Show(Ex.ToString(), "Error en Count de Registros");
                    }

                    if (count == 0)
                    {

                        switch (Comp)
                        {
                                /*
                            case "Qualitas":
                                this.workerThread = null;
                                this.workerThread = new Thread(
                                        new ThreadStart(this.qualitasToStand)
                                    );
                                this.workerThread.Start();
                                break;
                                  */

                            case "Afirme":
                                this.workerThread = null;
                                this.workerThread = new Thread(
                                        new ThreadStart(this.afirmeToStand)
                                    );
                                this.workerThread.Start();
                                break;

                            default:
                                //MessageBox.Show("Compañía no disponible", "Error");
                                int nComp = Convert.ToInt32(txtCompID2.Text);
                                this.workerThread = null;
                                this.workerThread = new Thread(
                                    () => DBToStand(Comp, nComp)
                                );
                                this.workerThread.Start();
                                break;
                        }
                    }
                    else {
                        MessageBox.Show("La tabla ya tiene datos!", "Error");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), "Error");
                }
            }
            else {
                MessageBox.Show("Selecciona una Compañía", "Error");
            }
            cmbCompany2.Enabled = true;
        }

        private void btn_Delete0_Click(object sender, EventArgs e)
        {
            if (txtOld0.Text != "" && cmbCompany0.SelectedIndex >= 0)
            {
                DialogResult dialogResult = MessageBox.Show("¿Borrar contenido especificado?", "Confirmación", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    if (cmbField0.SelectedIndex < 0) {
                        deletStringAll(" " + txtOld0.Text + " ", cmbCompany0.SelectedItem.ToString(), "DescripSimple");
                    } else {
                        MessageBox.Show("Not Available Yet", "ERROR");
                    }
                }
            }
            else {
                MessageBox.Show("You eat it at deleting", "ERROR");
            }

        }
    }
}