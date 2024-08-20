using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Windows.Forms;
using static ClosedXML.Excel.XLPredefinedFormat;

namespace VerificadorCoincidencias
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        //pegar click derecho dgv1 
        private void pegarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string textoCopiado = "", ultimo_vacio = "";
            int i = 0;

            if (Clipboard.ContainsData(DataFormats.Text))
            {
                textoCopiado = Clipboard.GetText();

                dataGridView1.DataSource = null;
                dataGridView1.Columns.Clear();
                dataGridView1.Rows.Clear();
                dataGridView1.DataSource = numerarDataTable(adaptarDataTable(textoCopiado, 1));
                dataGridView1.Columns["#"].Width = 50;

                i = dataGridView1.Rows.Count - 1;
                while (ultimo_vacio.Length == 0)
                {
                    ultimo_vacio = dataGridView1[1, i].Value.ToString();

                    if (ultimo_vacio.Length == 0)
                    {
                        dataGridView1.Rows.RemoveAt(i);
                    }

                    i--;
                }

                label4.Text = "Registros: " + dataGridView1.Rows.Count;

                if (dataGridView1.Rows.Count > 0)
                {
                    dataGridView2.Enabled = true;
                    dataGridView2.Focus();
                }
            }
        }
        //pegar click derecho dgv2 
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string textoCopiado = "", ultimo_vacio = "";
            int i = 0;

            if (Clipboard.ContainsData(DataFormats.Text))
            {
                textoCopiado = Clipboard.GetText();

                dataGridView2.DataSource = null;
                dataGridView2.Columns.Clear();
                dataGridView2.Rows.Clear();
                dataGridView2.DataSource = numerarDataTable(adaptarDataTable(textoCopiado, 2));
                dataGridView2.Columns["#"].Width = 50;

                i = dataGridView2.Rows.Count - 1;
                while (ultimo_vacio.Length == 0)
                {
                    ultimo_vacio = dataGridView2[1, i].Value.ToString();

                    if (ultimo_vacio.Length == 0)
                    {
                        dataGridView2.Rows.RemoveAt(i);
                    }

                    i--;
                }

                label5.Text = "Registros: " + dataGridView2.Rows.Count;

                if (dataGridView2.Rows.Count > 0)
                {
                    dataGridView3.Enabled = true;
                    button1.Enabled = true;
                    button6.Enabled = true;
                    button1.Focus();
                }
            }
        }
        //pegar click derecho dgv3
        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            string textoCopiado = "", ultimo_vacio = "";
            int i = 0;

            if (Clipboard.ContainsData(DataFormats.Text))
            {
                textoCopiado = Clipboard.GetText();

                dataGridView3.DataSource = null;
                dataGridView3.Columns.Clear();
                dataGridView3.Rows.Clear();
                dataGridView3.DataSource = numerarDataTable(adaptarDataTable(textoCopiado, 3));

                i = dataGridView3.Rows.Count - 1;
                while (ultimo_vacio.Length == 0)
                {
                    ultimo_vacio = dataGridView3[0, i].Value.ToString();

                    if (ultimo_vacio.Length == 0)
                    {
                        dataGridView3.Rows.RemoveAt(i);
                    }

                    i--;
                }

                label10.Text = "Registros: " + dataGridView3.Rows.Count;
            }
        }

        private System.Data.DataTable adaptarDataTable(string datos, int num_grid)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            int totfilas = 0, totColumnas = 0, tipo_col_apell = 0, tipo_mins = 0;
            string letra = "";

            for (int i = 0; i < datos.Length; i++)
            {
                letra = datos[i].ToString();

                if (letra == "\n")
                {
                    totfilas++;
                }

                if (letra == "\t")
                {
                    totColumnas++;
                }
            }

            string[] listLineas = new string[totfilas + 1];
            string[] linea = new string[totColumnas + 1];

            listLineas = datos.Split("\r\n");
            linea = listLineas[0].Split("\t");

            //Obtener tipo col Apellidos
            if (radioButton1.Checked)
            {
                tipo_col_apell = 2;
            }
            else
            {
                if (radioButton2.Checked)
                {
                    tipo_col_apell = 1;
                }
            }

            //Obtener tipo col minutos
            if (checkBox1.Checked)
            {
                tipo_mins = 1;
            }
            else
            {
                tipo_mins = 0;
            }

            if (num_grid == 1)
            {
                if (tipo_col_apell == 1)
                {
                    for (int i = 0; i < linea.Length; i++)
                    {

                        if (i == 0)
                        {
                            dt.Columns.Add("Nombre(s)");
                        }

                        if (i == 1)
                        {
                            dt.Columns.Add("Apellido(s)");
                        }

                        if (i == 2)
                        {
                            dt.Columns.Add("Correo-E");
                        }

                        if (i > 2)
                        {
                            dt.Columns.Add("COLUMNA_" + (i + 1));
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < linea.Length; i++)
                    {

                        if (i == 0)
                        {
                            dt.Columns.Add("Nombre(s)");
                        }

                        if (i == 1)
                        {
                            dt.Columns.Add("Primer Apellido");
                        }

                        if (i == 2)
                        {
                            dt.Columns.Add("Segundo Apellido");
                        }

                        if (i == 3)
                        {
                            dt.Columns.Add("Correo-E");
                        }

                        if (i > 3)
                        {
                            dt.Columns.Add("COLUMNA_" + (i + 1));
                        }
                    }
                }
            }

            if (num_grid == 2)
            {
                for (int i = 0; i < linea.Length; i++)
                {
                    if (tipo_mins == 0)
                    {
                        if (i == 0)
                        {
                            dt.Columns.Add("Nombre Completo");
                        }

                        if (i == 1)
                        {
                            dt.Columns.Add("Correo-E");
                        }

                        if (i > 1)
                        {
                            dt.Columns.Add("COLUMNA_" + (i + 1));
                        }
                    }
                    else
                    {
                        if (i == 0)
                        {
                            dt.Columns.Add("Nombre Completo");
                        }

                        if (i == 1)
                        {
                            dt.Columns.Add("Correo-E");
                        }

                        if (i == 2)
                        {
                            dt.Columns.Add("Minutos Totales");
                        }

                        if (i > 2)
                        {
                            dt.Columns.Add("COLUMNA_" + (i + 1));
                        }
                    }

                }
            }

            if (num_grid == 3)
            {
                for (int i = 0; i < linea.Length; i++)
                {
                    dt.Columns.Add("COLUMNA_" + (i + 1));
                }
            }

            for (int i = 0; i < listLineas.Length; i++)
            {
                linea = listLineas[i].Split("\t");
                dt.Rows.Add(linea);
            }

            return dt;
        }

        private System.Data.DataTable numerarDataTable(System.Data.DataTable original)
        {

            System.Data.DataTable dt = new System.Data.DataTable();
            int num = 0;

            dt.Columns.Add("#", num.GetType());

            for (int i = 0; i < original.Columns.Count; i++)
            {
                dt.Columns.Add(original.Columns[i].ColumnName);
            }

            for (int i = 0; i < original.Rows.Count; i++)
            {
                dt.Rows.Add();
                dt.Rows[i][0] = i + 1;

                for (int j = 0; j < original.Columns.Count; j++)
                {
                    dt.Rows[i][j + 1] = original.Rows[i][j];
                }
            }

            return dt;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            proceso_verificar();
        }

        private void proceso_verificar()
        {
            string aux = "", correoE = "", tipo_verif = "", variante1 = "", variante2 = "", variante3 = "", variante4 = "", variante5 = "", variante6 = "", num_mins = "0";
            int coincide = 0, coincideCompara = 0, tot_coincidencia = 0, segu = 0, tipo_col_apell = 0, tipo_mins = 0, minimo_asis = 0, tot_asis = 0;

            string[] nombreDividido = new string[5];
            string[] apellidoDividido = new string[5];

            System.Data.DataTable tablaBase = null;
            System.Data.DataTable tablaCompara = null;
            System.Data.DataTable tablaResultado = null;

            tablaBase = dataGridView1.DataSource as System.Data.DataTable;
            tablaCompara = dataGridView2.DataSource as System.Data.DataTable;

            tablaResultado = tablaBase.Copy();

            //Obtener tipo col Apellidos
            if (radioButton1.Checked)
            {
                tipo_col_apell = 2;
            }
            else
            {
                if (radioButton2.Checked)
                {
                    tipo_col_apell = 1;
                }
            }

            //Obtener tipo col minutos
            if (checkBox1.Checked)
            {
                tipo_mins = 1;
                minimo_asis = int.Parse(numericUpDown1.Value.ToString());
            }
            else
            {
                tipo_mins = 0;
            }

            //estandarizar textos
            for (int i = 0; i < tablaResultado.Rows.Count; i++)
            {
                for (int j = 0; j < tablaResultado.Columns.Count; j++)
                {
                    aux = "";

                    aux = normalizar_texto(tablaResultado.Rows[i][j].ToString());
                    tablaResultado.Rows[i][j] = aux;
                }
            }

            for (int i = 0; i < tablaCompara.Rows.Count; i++)
            {
                for (int j = 0; j < tablaCompara.Columns.Count; j++)
                {
                    aux = "";

                    aux = normalizar_texto(tablaCompara.Rows[i][j].ToString());
                    tablaCompara.Rows[i][j] = aux;
                }
            }

            if (tablaResultado.Columns.Contains("ASISTENCIA") == false)
            {
                tablaResultado.Columns.Add("ASISTENCIA");
            }

            if (tipo_mins == 1)
            {
                if (tablaResultado.Columns.Contains("ASISTENCIA_TIEMPO") == false)
                {
                    tablaResultado.Columns.Add("ASISTENCIA_TIEMPO");
                }
            }

            if (tablaResultado.Columns.Contains("TIPO_VERIFICACION") == false)
            {
                tablaResultado.Columns.Add("TIPO_VERIFICACION");
            }

            if (tablaCompara.Columns.Contains("ENCONTRADO") == false)
            {
                tablaCompara.Columns.Add("ENCONTRADO");
            }

            if (tablaCompara.Columns.Contains("TIPO_VERIFICACION") == false)
            {
                tablaCompara.Columns.Add("TIPO_VERIFICACION");
            }

            if (tipo_mins == 1)
            {
                if (tablaResultado.Columns.Contains("Minutos Totales") == false)
                {
                    tablaResultado.Columns.Add("Minutos Totales");
                }

                for (int i = 0; i < tablaResultado.Rows.Count; i++)
                {
                    tablaResultado.Rows[i]["Minutos Totales"] = "0";
                    tablaResultado.Rows[i]["ASISTENCIA_TIEMPO"] = "0";
                }
            }

            for (int i = 0; i < tablaCompara.Rows.Count; i++)
            {
                tablaCompara.Rows[i]["ENCONTRADO"] = "0";
            }

            tot_coincidencia = 0;
            tot_asis = 0;

            //buscar coincidencias
            for (int i = 0; i < tablaResultado.Rows.Count; i++)
            {
                aux = "";
                segu = 0;
                num_mins = "0";
                nombreDividido = null;
                apellidoDividido = null;
                correoE = "";
                variante1 = "";
                variante2 = "";
                variante3 = "";
                variante4 = "";
                variante5 = "";
                variante6 = "";

                //obtener valores linea
                for (int j = 0; j < tablaResultado.Columns.Count; j++)
                {
                    aux = tablaResultado.Rows[i][j].ToString();

                    if (tipo_col_apell == 1)
                    {
                        if (j == 1)
                        {
                            nombreDividido = aux.Split(' ');
                        }

                        if (j == 2)
                        {
                            apellidoDividido = aux.Split(' ');
                        }

                        if (j == 3)
                        {
                            correoE = aux;
                        }
                    }

                    if (tipo_col_apell == 2)
                    {
                        if (j == 1)
                        {
                            nombreDividido = aux.Split(' ');
                        }

                        if (j == 2)
                        {
                            apellidoDividido = aux.Split(' ');
                        }

                        if (j == 4)
                        {
                            correoE = aux;
                        }
                    }
                }

                //obtener nombre separado
                for (int l = 0; l < nombreDividido.Length; l++)
                {
                    if (nombreDividido[l].Length > 0)
                    {
                        variante1 += nombreDividido[l] + " "; //nombre completo
                        variante6 += nombreDividido[l] + " "; //nombre completo

                        if (l == 0)
                        {
                            variante2 = nombreDividido[l] + " ";//primer nombre
                            variante3 = nombreDividido[l] + " ";//primer nombre
                        }

                        if (l > 0 && segu == 0)
                        {
                            if (nombreDividido[l].Equals("de") || nombreDividido[l].Equals("del") || nombreDividido[l].Equals("la"))
                            {
                            }
                            else
                            {
                                variante4 = nombreDividido[l] + " ";//segundo nombre
                                variante5 = nombreDividido[l] + " ";//segundo nombre
                                segu = 1;
                            }
                        }
                    }
                }

                //obtener apellido separado
                for (int m = 0; m < apellidoDividido.Length; m++)
                {
                    if (apellidoDividido[m].Length > 0)
                    {
                        variante1 += apellidoDividido[m] + " ";//nombre completo
                        variante3 += apellidoDividido[m] + " ";//apellidos completos
                        variante5 += apellidoDividido[m] + " ";//apellidos completos

                        if (tipo_col_apell == 1)
                        {
                            if (m == 0)
                            {
                                variante2 += apellidoDividido[m];//primer apellido
                                variante4 += apellidoDividido[m];//primer apellido
                                variante6 += apellidoDividido[m];//primer apellido
                            }
                        }
                        else
                        {
                            if (tipo_col_apell == 2)
                            {
                                variante2 += apellidoDividido[m];//primer apellido
                                variante4 += apellidoDividido[m];//primer apellido
                                variante6 += apellidoDividido[m];//primer apellido
                            }

                        }
                    }
                }

                variante1 = variante1.TrimEnd(' ');
                variante2 = variante2.TrimEnd(' ');
                variante3 = variante3.TrimEnd(' ');
                variante4 = variante4.TrimEnd(' ');
                variante5 = variante5.TrimEnd(' ');
                variante6 = variante6.TrimEnd(' ');

                coincide = 0;
                coincideCompara = 0;

                //Primera Busqueda: Nombre y Apellidos completos
                tipo_verif = "Nombres y Apellidos Completos";
                for (int k = 0; k < tablaCompara.Rows.Count; k++)
                {
                    if (tablaCompara.Rows[k]["ENCONTRADO"].ToString() != "1")
                    {
                        //coincide nombre y apellidos completos
                        if (variante1.Equals(tablaCompara.Rows[k]["Nombre Completo"].ToString()))
                        {
                            coincide = 1;
                            coincideCompara = 1;
                        }

                        if (coincideCompara == 1)
                        {
                            tablaCompara.Rows[k]["ENCONTRADO"] = coincideCompara.ToString();
                            tablaCompara.Rows[k]["TIPO_VERIFICACION"] = tipo_verif;
                            coincideCompara = 0;

                            if (tipo_mins == 1)
                            {
                                num_mins = tablaCompara.Rows[k]["Minutos Totales"].ToString();
                            }
                        }
                    }
                }

                if (coincide == 0)
                {
                    // Segunda Busqueda: CorreoE
                    tipo_verif = "Correo Electronico";
                    for (int k = 0; k < tablaCompara.Rows.Count; k++)
                    {
                        if (tablaCompara.Rows[k]["ENCONTRADO"].ToString() != "1")
                        {
                            //coincide correoE
                            if (correoE.Equals(tablaCompara.Rows[k]["Correo-E"].ToString()))
                            {
                                coincide = 1;
                                coincideCompara = 1;
                            }

                            if (coincideCompara == 1)
                            {
                                tablaCompara.Rows[k]["ENCONTRADO"] = coincideCompara.ToString();
                                tablaCompara.Rows[k]["TIPO_VERIFICACION"] = tipo_verif;
                                coincideCompara = 0;

                                if (tipo_mins == 1)
                                {
                                    num_mins = tablaCompara.Rows[k]["Minutos Totales"].ToString();
                                }
                            }
                        }
                    }
                }

                if (coincide == 0)
                {
                    //Tercera Busqueda: Nombre y Apellidos Completos
                    tipo_verif = "Primer Nombre y Apellidos Completos";
                    for (int k = 0; k < tablaCompara.Rows.Count; k++)
                    {
                        if (tablaCompara.Rows[k]["ENCONTRADO"].ToString() != "1")
                        {
                            //coincide primer nombre y apellidos completos
                            if (variante3.Equals(tablaCompara.Rows[k]["Nombre Completo"].ToString()))
                            {
                                coincide = 1;
                                coincideCompara = 1;
                            }

                            if (coincideCompara == 1)
                            {
                                tablaCompara.Rows[k]["ENCONTRADO"] = coincideCompara.ToString();
                                tablaCompara.Rows[k]["TIPO_VERIFICACION"] = tipo_verif;
                                coincideCompara = 0;

                                if (tipo_mins == 1)
                                {
                                    num_mins = tablaCompara.Rows[k]["Minutos Totales"].ToString();
                                }
                            }
                        }
                    }
                }

                if (coincide == 0)
                {
                    //Cuarta Busqueda: Segundo Nombre y Apellidos Completos
                    tipo_verif = "Segundo Nombre y Apellidos Completos";
                    for (int k = 0; k < tablaCompara.Rows.Count; k++)
                    {
                        if (tablaCompara.Rows[k]["ENCONTRADO"].ToString() != "1")
                        {
                            //coincide segundo nombre y apellidos completos
                            if (variante5.Equals(tablaCompara.Rows[k]["Nombre Completo"].ToString()))
                            {
                                coincide = 1;
                                coincideCompara = 1;
                            }

                            if (coincideCompara == 1)
                            {
                                tablaCompara.Rows[k]["ENCONTRADO"] = coincideCompara.ToString();
                                tablaCompara.Rows[k]["TIPO_VERIFICACION"] = tipo_verif;
                                coincideCompara = 0;

                                if (tipo_mins == 1)
                                {
                                    num_mins = tablaCompara.Rows[k]["Minutos Totales"].ToString();
                                }
                            }
                        }
                    }
                }

                if (coincide == 0)
                {
                    //Quinta Busqueda: Primer Nombre y Primer Apellido
                    tipo_verif = "Primer Nombre y Primer Apellido";
                    for (int k = 0; k < tablaCompara.Rows.Count; k++)
                    {
                        if (tablaCompara.Rows[k]["ENCONTRADO"].ToString() != "1")
                        {
                            //coincide primer nombre y primer apellido
                            if (variante2.Equals(tablaCompara.Rows[k]["Nombre Completo"].ToString()))
                            {
                                coincide = 1;
                                coincideCompara = 1;
                            }

                            if (coincideCompara == 1)
                            {
                                tablaCompara.Rows[k]["ENCONTRADO"] = coincideCompara.ToString();
                                tablaCompara.Rows[k]["TIPO_VERIFICACION"] = tipo_verif;
                                coincideCompara = 0;

                                if (tipo_mins == 1)
                                {
                                    num_mins = tablaCompara.Rows[k]["Minutos Totales"].ToString();
                                }
                            }
                        }
                    }
                }

                if (coincide == 0)
                {
                    //Sexta Busqueda: Segundo Nombre y Primer Apellido
                    tipo_verif = "Segundo Nombre y Primer Apellido";
                    for (int k = 0; k < tablaCompara.Rows.Count; k++)
                    {
                        if (tablaCompara.Rows[k]["ENCONTRADO"].ToString() != "1")
                        {
                            //coincide segundo nombre y primer apellido 
                            if (variante4.Equals(tablaCompara.Rows[k]["Nombre Completo"].ToString()))
                            {
                                coincide = 1;
                                coincideCompara = 1;
                            }

                            if (coincideCompara == 1)
                            {
                                tablaCompara.Rows[k]["ENCONTRADO"] = coincideCompara.ToString();
                                tablaCompara.Rows[k]["TIPO_VERIFICACION"] = tipo_verif;
                                coincideCompara = 0;

                                if (tipo_mins == 1)
                                {
                                    num_mins = tablaCompara.Rows[k]["Minutos Totales"].ToString();
                                }
                            }
                        }
                    }
                }

                /*
                //variantes
                for (int k = 0; k < tablaCompara.Rows.Count; k++)
                {
                    if (tablaCompara.Rows[k][1].ToString() != "1")
                    {
                        //coincide nombre y apellidos completos
                        if (variante1.Equals(tablaCompara.Rows[k][0].ToString()))
                        {
                            coincide = 1;
                            coincideCompara = 1;
                        }

                        //coincide primer nombre y apellidos completos
                        if (variante3.Equals(tablaCompara.Rows[k][0].ToString()))
                        {
                            coincide = 1;
                            coincideCompara = 1;
                        }

                        //coincide primer nombre y apellido
                        if (variante2.Equals(tablaCompara.Rows[k][0].ToString()))
                        {
                            coincide = 1;
                            coincideCompara = 1;
                        }

                        //coincide segundo nombre y primer apellido 
                        if (variante4.Equals(tablaCompara.Rows[k][0].ToString()))
                        {
                            coincide = 1;
                            coincideCompara = 1;
                        }

                        //coincide segundo nombre y apellidos completos
                        if (variante5.Equals(tablaCompara.Rows[k][0].ToString()))
                        {
                            coincide = 1;
                            coincideCompara = 1;
                        }

                        //coincide nombres completos y primer apellido
                        if (variante6.Equals(tablaCompara.Rows[k][0].ToString()))
                        {
                            coincide = 1;
                            coincideCompara = 1;
                        }

                        if (coincideCompara == 1)
                        {
                            tablaCompara.Rows[k][1] = coincideCompara.ToString();
                            coincideCompara = 0;
                        }
                    }
                }
                */


                if (coincide == 1)
                {
                    tablaResultado.Rows[i]["TIPO_VERIFICACION"] = tipo_verif;

                    if (tipo_mins == 1)
                    {
                        tablaResultado.Rows[i]["Minutos Totales"] = num_mins;
                        int tot_min_asis = 0;

                        if (int.TryParse(num_mins, out tot_min_asis))
                        {
                            if (tot_min_asis >= minimo_asis)
                            {
                                tot_asis++;
                                tablaResultado.Rows[i]["ASISTENCIA_TIEMPO"] = 1;
                            }
                            else
                            {
                                tablaResultado.Rows[i]["ASISTENCIA_TIEMPO"] = 0;
                            }
                        }


                    }

                    tot_coincidencia++;
                }
                else
                {
                    tablaResultado.Rows[i]["TIPO_VERIFICACION"] = "-";
                }

                tablaResultado.Rows[i]["ASISTENCIA"] = coincide.ToString();
            }

            dataGridView3.DataSource = null;
            dataGridView3.Columns.Clear();
            dataGridView3.Rows.Clear();
            dataGridView3.DataSource = tablaResultado;
            dataGridView3.Columns["#"].Width = 50;

            label7.Text = tot_coincidencia.ToString();
            label9.Text = tablaResultado.Rows.Count.ToString();
            label10.Text = "Registros: " + dataGridView3.RowCount;

            if (tipo_mins == 1)
            {
                MessageBox.Show("Comparación Terminada\n" + tot_coincidencia + "/" + tablaResultado.Rows.Count + " Asistencias Encontradas\n" + tot_asis + " Asistencias de acuerdo al tiempo Mínimo", "ATENCION", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Comparación Terminada\n" + tot_coincidencia + "/" + tablaResultado.Rows.Count + " Asistencias Encontradas", "ATENCION", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private string normalizar_texto(string texto)
        {
            string resultado = "";
            string[] palabras = new string[10];

            resultado = texto.ToLower();

            if (resultado.Contains('á'))
            {
                resultado = resultado.Replace('á', 'a');
            }

            if (resultado.Contains('é'))
            {
                resultado = resultado.Replace('é', 'e');
            }

            if (resultado.Contains('í'))
            {
                resultado = resultado.Replace('í', 'i');
            }

            if (resultado.Contains('ó'))
            {
                resultado = resultado.Replace('ó', 'o');
            }

            if (resultado.Contains('ú'))
            {
                resultado = resultado.Replace('ú', 'u');
            }

            palabras = resultado.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            resultado = "";

            for (int i = 0; i < palabras.Length; i++)
            {
                resultado += palabras[i] + " ";
            }

            resultado = resultado.TrimStart();
            resultado = resultado.TrimEnd();

            return resultado;
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (dataGridView3.Rows.Count > 0)
            {
                string nom_arch = "";
                System.Data.DataTable tablaExportar = null;
                System.Data.DataTable tablaExportar2 = null;

                tablaExportar = dataGridView3.DataSource as System.Data.DataTable;

                if (dataGridView2.Rows.Count > 0)
                {

                    tablaExportar2 = dataGridView2.DataSource as System.Data.DataTable;
                }

                SaveFileDialog dialog_save = new SaveFileDialog();
                dialog_save.Filter = "Archivos de Excel (*.XLSX)|*.XLSX"; //le indicamos el tipo de filtro en este caso que busque solo los archivos excel

                dialog_save.Title = "Guardar Listas de Resultados";//le damos un titulo a la ventana

                if (dialog_save.ShowDialog() == DialogResult.OK)
                {
                    nom_arch = dialog_save.FileName; //open file
                    XLWorkbook wb = new XLWorkbook();
                    wb.Worksheets.Add(tablaExportar, "Asistencias");

                    if (dataGridView2.Rows.Count > 0)
                    {
                        wb.Worksheets.Add(tablaExportar2, "ListaReporte");
                    }

                    wb.SaveAs(nom_arch);

                    MessageBox.Show("El archivo se generó correctamente.", "¡Exito!");
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.Focus();
        }

        private string copiarTablaPortapapeles(int numDataGrid)
        {
            string textoRespuesta = "", linea = "";

            System.Data.DataTable dataTable = dataTable = new System.Data.DataTable();

            if (numDataGrid == 2)
            {
                dataTable = dataGridView2.DataSource as System.Data.DataTable;
            }
            else if (numDataGrid == 3)
            {
                dataTable = dataGridView3.DataSource as System.Data.DataTable;
            }

            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                linea += dataTable.Columns[i].ColumnName + "\t";
            }

            textoRespuesta += linea + "\r\n";

            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                linea = "";

                for (int j = 0; j < dataTable.Columns.Count; j++)
                {
                    linea += dataTable.Rows[i][j].ToString() + "\t";
                }

                textoRespuesta += linea + "\r\n";
            }

            return textoRespuesta;
        }

        //copiar
        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            if (dataGridView2.ContainsFocus == true)
            {
                if (dataGridView2.Rows.Count > 0)
                {
                    Clipboard.SetText(copiarTablaPortapapeles(2));
                }
            }
            else
            {

                if (dataGridView3.ContainsFocus == true)
                {
                    if (dataGridView2.Rows.Count > 0)
                    {
                        Clipboard.SetText(copiarTablaPortapapeles(3));
                    }
                }
            }
        }
        //pegar
        private void toolStripMenuItem6_Click(object sender, EventArgs e)
        {
            string textoCopiado = "", ultimo_vacio = "";
            int i = 0;

            if (dataGridView1.ContainsFocus == true)
            {
                if (Clipboard.ContainsData(DataFormats.Text))
                {
                    textoCopiado = Clipboard.GetText();

                    dataGridView1.DataSource = null;
                    dataGridView1.Columns.Clear();
                    dataGridView1.Rows.Clear();
                    dataGridView1.DataSource = numerarDataTable(adaptarDataTable(textoCopiado, 1));
                    dataGridView1.Columns["#"].Width = 50;

                    i = dataGridView1.Rows.Count - 1;
                    while (ultimo_vacio.Length == 0)
                    {
                        ultimo_vacio = dataGridView1[1, i].Value.ToString();

                        if (ultimo_vacio.Length == 0)
                        {
                            dataGridView1.Rows.RemoveAt(i);
                        }

                        i--;
                    }

                    label4.Text = "Registros: " + dataGridView1.Rows.Count;

                    if (dataGridView1.Rows.Count > 0)
                    {
                        dataGridView2.Enabled = true;
                        dataGridView2.Focus();
                    }
                }
            }
            else
            {

                if (dataGridView2.ContainsFocus == true)
                {
                    if (Clipboard.ContainsData(DataFormats.Text))
                    {
                        textoCopiado = Clipboard.GetText();

                        dataGridView2.DataSource = null;
                        dataGridView2.Columns.Clear();
                        dataGridView2.Rows.Clear();
                        dataGridView2.DataSource = numerarDataTable(adaptarDataTable(textoCopiado, 2));
                        dataGridView2.Columns["#"].Width = 50;

                        i = dataGridView2.Rows.Count - 1;
                        while (ultimo_vacio.Length == 0)
                        {
                            ultimo_vacio = dataGridView2[1, i].Value.ToString();

                            if (ultimo_vacio.Length == 0)
                            {
                                dataGridView2.Rows.RemoveAt(i);
                            }

                            i--;
                        }

                        label5.Text = "Registros: " + dataGridView2.Rows.Count;

                        if (dataGridView2.Rows.Count > 0)
                        {
                            dataGridView3.Enabled = true;
                            button1.Enabled = true;
                            button6.Enabled = true;
                            button6.Focus();
                        }
                    }
                }
            }

            /*
            if (dataGridView3.ContainsFocus == true)
            {
                if (Clipboard.ContainsData(DataFormats.Text))
                {
                    textoCopiado = Clipboard.GetText();

                    dataGridView3.DataSource = null;
                    dataGridView3.Columns.Clear();
                    dataGridView3.Rows.Clear();
                    dataGridView3.DataSource = adaptarDataTable(textoCopiado, 3);

                    i = dataGridView3.Rows.Count - 1;
                    while (ultimo_vacio.Length == 0)
                    {
                        ultimo_vacio = dataGridView3[0, i].Value.ToString();

                        if (ultimo_vacio.Length == 0)
                        {
                            dataGridView3.Rows.RemoveAt(i);
                        }

                        i--;
                    }

                    label10.Text = "Registros: " + dataGridView3.Rows.Count;
                }
            }
        */
        }

        private void dataGridView1_Enter(object sender, EventArgs e)
        {
            dataGridView1.BackgroundColor = SystemColors.ControlDarkDark;
            dataGridView2.BorderStyle = BorderStyle.Fixed3D;
        }

        private void dataGridView1_Leave(object sender, EventArgs e)
        {
            dataGridView1.BackgroundColor = SystemColors.ControlDark;
            dataGridView2.BorderStyle = BorderStyle.FixedSingle;
        }

        private void dataGridView2_Enter(object sender, EventArgs e)
        {
            dataGridView2.BackgroundColor = SystemColors.ControlDarkDark;
            dataGridView2.BorderStyle = BorderStyle.Fixed3D;
        }

        private void dataGridView2_Leave(object sender, EventArgs e)
        {
            dataGridView2.BackgroundColor = SystemColors.ControlDark;
            dataGridView2.BorderStyle = BorderStyle.FixedSingle;
        }

        private void dataGridView3_Enter(object sender, EventArgs e)
        {
            dataGridView3.BackgroundColor = SystemColors.ControlDarkDark;
            dataGridView3.BorderStyle = BorderStyle.Fixed3D;
        }

        private void dataGridView3_Leave(object sender, EventArgs e)
        {
            dataGridView3.BackgroundColor = SystemColors.ControlDark;
            dataGridView3.BorderStyle = BorderStyle.FixedSingle;
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            if (dataGridView2.Rows.Count > 0)
            {
                Clipboard.SetText(copiarTablaPortapapeles(2));
            }
        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            if (dataGridView3.Rows.Count > 0)
            {
                Clipboard.SetText(copiarTablaPortapapeles(3));
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            limpiar_dgv1();
            separados_dgv1();

            limpiar_dgv2();

            dataGridView3.DataSource = null;
            dataGridView3.Columns.Clear();
            dataGridView3.Rows.Clear();
            label10.Text = "Registros: 0";
            dataGridView3.Enabled = false;

            button1.Enabled = false;
            button6.Enabled = false;

            label7.Text = "???";
            label9.Text = "???";

            dataGridView3.Columns.Add("num", "#");
            dataGridView3.Columns["num"].Width = 50;

            radioButton1.Checked = true;
            radioButton2.Checked = false;

            checkBox1.Checked = false;
            numericUpDown1.Value = 0;
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            limpiar_dgv1();
            separados_dgv1();
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            limpiar_dgv1();
            combinados_dgv1();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            limpiar_dgv2();

            if (checkBox1.Checked == true)
            {
                numericUpDown1.Enabled = true;
                con_minutos();
            }
            else
            {
                numericUpDown1.Enabled = false;
                sin_minutos();
            }
        }

        private void limpiar_dgv1()
        {
            dataGridView1.DataSource = null;
            dataGridView1.Columns.Clear();
            dataGridView1.Rows.Clear();
            label4.Text = "Registros: 0";
            dataGridView1.Focus();
        }

        private void separados_dgv1()
        {
            dataGridView1.Columns.Add("num", "#");
            dataGridView1.Columns["num"].Width = 50;
            dataGridView1.Columns.Add("nombres", "Nombre(s)");
            dataGridView1.Columns.Add("pri_apellido", "Primer Apellido");
            dataGridView1.Columns.Add("seg_apellido", "Segundo Apellido");
            dataGridView1.Columns.Add("correo", "Correo-E");
        }

        private void combinados_dgv1()
        {
            dataGridView1.Columns.Add("num", "#");
            dataGridView1.Columns["num"].Width = 50;
            dataGridView1.Columns.Add("nombres", "Nombre(s)");
            dataGridView1.Columns.Add("apellidos", "Apellido(s)");
            dataGridView1.Columns.Add("correo", "Correo-E");
        }

        private void limpiar_dgv2()
        {
            dataGridView2.DataSource = null;
            dataGridView2.Columns.Clear();
            dataGridView2.Rows.Clear();
            label5.Text = "Registros: 0";
            //dataGridView2.Enabled = false;
        }

        private void sin_minutos()
        {
            dataGridView2.Columns.Add("num", "#");
            dataGridView2.Columns["num"].Width = 50;
            dataGridView2.Columns.Add("nombres", "Nombre Completo");
            dataGridView2.Columns.Add("correo", "Correo-E");
        }

        private void con_minutos()
        {
            dataGridView2.Columns.Add("num", "#");
            dataGridView2.Columns["num"].Width = 50;
            dataGridView2.Columns.Add("nombres", "Nombre Completo");
            dataGridView2.Columns.Add("correo", "Correo-E");
            dataGridView2.Columns.Add("minutos", "Minutos Totales");
        }

    }
}