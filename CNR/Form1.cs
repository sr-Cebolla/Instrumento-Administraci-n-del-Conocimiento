using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;

namespace CNR
{
    public partial class Form1 : Form
    {
        StreamWriter Wrt;
        HGAdminC a = new HGAdminC();
        int i = 0;
        int uno = 0, dos = 0, tres = 0, cuatro = 0, cinco = 0;
        string guardado = "", temp1 = "", temp2 = "", temp3 = "", temp4 = "", temp5 = "", error = "", errorCom = "";
        bool pollo1 = false, pollo2 = false, pollo3 = true, pollo4 = true, pollo5 = true;

        /// <summary>
        /// Informacion de calificacion (conteo de puntos)
        /// </summary>
        string[,] SDC = {
            {"Excelente", "Supera el conocimiento requerido.", "5"},
            {"Muy Bueno","Domina el conocimiento requerido.","4"},
            {"Bueno","Alcanza el conocimiento requerido.","3"},
            {"Regular","Esta proximo a alcanzar el conocimiento requerido.","2"},
            {"Deficiente","No alcanza el conocimiento requerido.","1"}
        };
        /// <summary>
        /// Herramienta gerencial de administracion del conocimiento
        /// </summary>
        string[,] SDP1 = {
                {"A.","Prevision"},//0
            {"1.", "¿Como califica la toma de medidas que aplica, para asegurar la eficiencia operativa del personal de su unidad?"},
            {"2.", "¿Cuan efectivo es su enfoque que utiliza para la toma de decisiones importantes?"},
            {"3.", "¿De que manera preve las actividades que ejecuta?"},
                {"B.","Planeacion"},//4
            {"4.","¿Como califica su habilidad para asignar recursos y asi alcanzar los objetivos trazados?"},
            {"5.","¿Como califica su pensamiento flexible e integral en el momento de planificar?"},
            {"6.","¿Cuan efectivos son los medios que utiliza para lograr los objetivos trazados?"},
                {"C.","Organizacion"},//8
            {"7.","¿Como califica su habilidad para fomentar el desarrollo profesional del personal de su unidad?"},
            {"8.","¿Como califica el nivel de reconocimiento del esfuerzo del trabajo por parte del personal de su unidad?"},
            {"9.","¿Cuan efectivo es el clima organizacional que brinda al personal de su unidad? "},
                {"D.","Integracion"},//12
            {"10.","¿Cuan efectivo son los metodos que emplea para evaluar el desempeño del personal de su unidad?"},
            {"11.","¿Como califica su empleo de estrategias para mejorar el desempeño del personal de su unidad?"},
            {"12.","¿Como califica la motivacion que les brinda al personal de su unidad en momentos dificiles?"},
                {"E.","Dirreccion"},//16
            {"13.","¿Cuan efectiva es la prioridad que aplica al realizar sus responsabilidades?"},
            {"14.","¿Cuan efectiva es su comunicacion respecto a la elaboracion de planes en su unidad?"},
            {"15.","¿Como califica el nivel de estrategias que aplica para crear un equipo de trabajo eficiente?"},
            {"16.","¿Cuan efectivas son las propuestas de mejora que brinda al personal de su unidad?"},
            {"17.","¿Cuan efectivas son las estrategias que ha utilizado para el manejo con clientes dificiles?"},
                {"F.","Control"},//22
            {"18.","¿Como califica su nivel de empleo de las politicas de control interno en su unidad?"},
            {"19.","¿Como califica su nivel de control sobre situaciones dificiles o urgentes presentadas en su unidad?"},
            {"20.","¿Como califica el enfoque que mantiene para lograr los objetivos?"}
        };
        /// <summary>
        /// Herramienta gerencial del conocimiento
        /// </summary>
        string[,] SDP2 = {
                {"A.","Formulacion de Planos"},//00
            {"1.", "¿Planes estrategicos alineados a mision, vision?"},
            {"2.", "¿Metas, objetivos especificos, alcanzables definidos?"},
            {"3.", "¿Planes detallados con actividades, responsables, cronogramas?"},
            {"4.","¿Revision, actualizacion periodica de aviones?"},
            {"5.","¿Evaluacion de riesgos, estrategias de mitigacion?"},
                {"B.","Control Administrativo Ejercido"},//06
            {"6.","¿Sistema de control, monitoreo implementado?"},
            {"7.","¿Indicadores clave de desempeño definidos?"},
            {"8.","¿Evaluaciones periodicas para identificar mejoras?"},
            {"9.","¿Acciones correctivas para asegurar el cumplimiento?"},
            {"10.","¿Fomento de mejora continua en procesos?"},
                {"C.","Manera de Dirigir"},//12
            {"11.","¿Liderazgo efectivo, inspirador establecido?"},
            {"12.","¿Ambiente positivo, trabajo en equipo promovido?"},
            {"13.","¿Estrategias de motivacion para personales implementadas?"},
            {"14.","¿Decisiones oportunas, basadas en informacion confiable?"},
            {"15.","¿Gestion eficiente de conflictos, problemas?"},
                {"D.","Comunicacion Efectiva"},//18
            {"16.","¿Canales de comunicacion efectivos establecidos?"},
            {"17.","¿Comunicacion abierta, valoracion de opiniones?"},
            {"18.","¿Actualizaciones, decisiones compartidas regularmente?"},
            {"19.","¿Diferentes canales utilizados, flujo de informacion?"},
            {"20.","¿Retroalimentacion constructiva del personal fomentado?"}
        };
        /// <summary>
        /// Herramienta gerencial de direccion del conocimiento
        /// </summary>
        string[,] SDP3 = {
                {"A.","Motivacion"},//00
            {"1.", "¿Como calificaria su habilidad para inspirar  a su equipo?"},
            {"2.", "¿Que tan efectivo es en valorar los logros de su equipo?"},
            {"3.", "¿Como calificaria su capacidad para fomentar el compromiso en su equipo?"},
                {"B.","Liderazgo"},//04
            {"4.","¿Que tan bien ejerce un liderazgo efectivo?"},
            {"5.","¿Como calificaria su habilidad para ser un modelo a seguir y generar confianza?"},
            {"6.","¿Que tan efectivo es en la toma de decisiones y resolucion de conflictos?"},
                {"C.","Comunicacion"},//08
            {"7.","¿Como calificaria su capacidad para comunicarse de manera clara?"},
            {"8.","¿Que tan bien fomenta un ambiente de comunicacion abierta?"},
            {"9.","¿Como calificaria su habilidad para escuchar diferentes perspectivas?"},
                {"D.","Desarrollo de Equipo"},//12
            {"10.","¿Que tan efectivo es en fomentar un ambiente de trabajo colaborativo?"},
            {"11.","¿Como calificaria su capacidad para delegar tareas de manera efectiva?"},
            {"12.","¿Que tan bien maneja el crecimiento profesional de su equipo?"},
                {"E.","Gestion de Cambio"},//16
            {"13.","¿Como calificaria su habilidad para liderar de manera efectiva?"},
            {"14.","¿Que tan bien comunica y genera aceptacion hacia los cambios necesarios?"},
            {"15.","¿Como calificaria su capacidad para adaptarse durante los cambios?"},
            {"16.","¿Que tan habil es para anticipar a su equipo ante los cambios venideros?"},
            {"17.","¿Como calificaria su capacidad para minimizar la resistencia de su equipo durante los procesos de cambio?"},
                {"F.","Vision Estrategica"},//22
            {"18.","¿Como calificaria su habilidad para establecer una vision clara con los objetivos organizacionales?"},
            {"19.","¿Que tan efectivo es en lograr que su equipo se comprometa con la vision?"},
            {"20.","¿Que tan efectivo es analizar el entorno?"}
        };
        /// <summary>
        /// Herramienta gerencial de liderazgo del conocimiento
        /// </summary>
        string[,] SDP4 = {
                {"A.","Democracia"},//00
            {"1.", "¿Como ha manejado situaciones de conflicto dentro de su equipo para garantizar que todas las voces sean escuchadas y consideradas?"},
            {"2.", "¿Como califica su capacidad para lograr equilibrar la toma de decisiones inclusiva con la necesidad de mantener la eficiencia y la direccion en su equipo?"},
            {"3.", "¿En que nivel considera que ha cultivado la confianza en sus relaciones con los miembros del equipo a traves de su practica de liderazgo democratico?"},
                {"B.","Participacion"},//04
            {"4.","¿Cual es el nivel de confianza que percibe en su equipo en relacion con la participacion en la toma de decisiones?"},
            {"5.","¿Cuanto tiempo dedica regularmente a facilitar reuniones o sesiones de lluvia de ideas donde se fomente la participacion activa de los miembros del equipo?"},
            {"6.","¿En que medida involucra a los miembros de su equipo en la generacion de ideas y soluciones para mejorar procesos?"},
                {"C.","Situacional"},//08
            {"7.","¿Con que frecuencia evalua la preparacion de sus seguidores para las diferentes tareas que enfrenta su equipo?"},
            {"8.","¿Como se ponderan sus habilidades para adaptarse mejor a las necesidades cambiantes de su equipo?"},
            {"9.","¿Como califica su facilidad de adaptabilidad a diferentes circunstancias?"},
                {"D.","Maternidad o Paternidad"},//12
            {"10.","¿Como califica el nivel de comunicacion comprensiva con los miembros del equipo?"},
            {"11.","¿Que tan efectivas han sido las relaciones cercanas que ha creado con los miembros de su equipo?"},
            {"12.","¿Cuanta es su empatia hacia los problemas que pasan su equipo de trabajo?"},
                {"E.","Contingencial"},//16
            {"13.","¿En que nivel considera que esta la madurez de su equipo para enfrentar problemas drasticos?"},
            {"14.","¿Que tan seguro/a esta de que cuando habla con su equipo sobre los objetivos de la empresa, ellos los comprendieron?"},
            {"15.","¿Como califica que tan efectivas han sido las estrategias que ha tenido que tomar para gestionar un buen cambio en situaciones contingentes?"},
            {"16.","¿Como califica su capacidad para establecer una cultura organizacional que valore la adaptacion al cambio?"},
            {"17.","¿Como calificaria la capacidad de abordar los desafios emocionales que pueden surgir en momentos de crisis?"},
                {"F.","Transformacional"},//22
            {"18.","¿Que tanto fomenta la innovacion dentro de su equipo de trabajo para que logren alcanzar las metas?"},
            {"19.","¿Con que frecuencia comunica una vision inspiradora y convincente para el futuro de su organizacion?"},
            {"20.","¿Como calificaria su capacidad para ejercer un liderazgo que motive a su equipo hacia el cambio?"}
        };
        /// <summary>
        /// Herramienta gerencial de gestion del conocimiento
        /// </summary>
        string[,] SDP5 = {
                {"A.","Red de apoyo"},//00
            {"1.", "¿Cuan efectivas son las estrategias que utiliza para ampliar su red de apoyo?"},
            {"2.", "¿Cuan efectivas son las  tacticas que utiliza,  para mantener las relaciones dentro de su red de apoyo?"},
            {"3.", "¿Como califica el uso efectivo de su red de apoyo en situaciones desafiantes?"},
            {"4.","¿Como califica la construccion de su red de apoyo?"},
                {"B.","Manera de Comunicarse"},//05
            {"5.","¿Como califica su manera de  comunicar la informacion que el personal de su unidad necesita comprender?"},
            {"6.","¿Cuan efectivos son los medios de comunicacion que utiliza para comunicar informacion relevante al personal de su unidad?"},
            {"7.","¿Como califica su manera de explicar al personal de su unidad las razones que sostuvo al tomar decisiones?"},
            {"8.","¿Como califica su nivel de compromiso que tiene con el  personal de su unidad para contribuir al desarrollo de los objetivos?"},
                {"C.","Forma de Delegar"},//10
            {"9.","¿Cuan efectiva es la division de responsabilidades que aplica en su unidad de trabajo ?"},
            {"10.","¿Como califica su compromiso al delegar las responsabilidades al personal de su unidad? "},
            {"11.","¿Como califica su capacidad de hacer que el personal de su unidad realice las responsabilidades delegadas?"},
            {"12.","¿Cuan efectivo es el analisis que realiza sobre las prioridades del personal de su unidad antes de delegar las responsabilidades?"},
                {"D.","Manera de Evaluacion"},//15
            {"13.","¿Como califica su habilidad para buscar las aportaciones del personal de su unidad?"},
            {"14.","¿En que medida maneja los problemas existentes sin crear otros nuevos?"},
            {"15.","¿Como califica el nivel que emplea para reconocer el trabajo que realiza el personal de su unidad?"},
                {"E.","Manera de Escucha Activa" },//19
            {"16.","¿Como califica su habilidad de escucha activa que brinda al personal de su unidad?"},
            {"17.","¿Cuan efectivas son las relaciones que construye cuando aplica la escucha activa ?"},
            {"18.","¿Cuan efectiva es la informacion que brinda al personal de su unidad al aplicar su escucha activa para resolver problemas?  "},
                {"F.","Generacion de Entorno Inclusivo"},//23
            {"19.","¿Como califica su habilidad de tratar de manera justa al personal de su unidad, independientemente de sus caracteristicas personales?"},
            {"20.","¿Cuan efectiva es su manera de fomentar la participacion activa del personal de su unidad para la colaboracion en la creacion de un ambiente inclusivo? "}
        };
        public Form1()
        {
            InitializeComponent();
            //tt DE AYUDA EN BOTONES
            ttHelp.SetToolTip(btnCalcDat1, "con este boton prodra calcular los resultados \ntomando en cuenta los resultados de la tabla 1");
            ttHelp.SetToolTip(btnCalcDat2, "con este boton prodra calcular los resultados \ntomando en cuenta los resultados de la tabla 2");
            ttHelp.SetToolTip(btnCalcDat3, "con este boton prodra calcular los resultados \ntomando en cuenta los resultados de la tabla 3");
            ttHelp.SetToolTip(btnCalcDat4, "con este boton prodra calcular los resultados \ntomando en cuenta los resultados de la tabla 4");
            ttHelp.SetToolTip(btnCalcDat5, "con este boton prodra calcular los resultados \ntomando en cuenta los resultados de la tabla 5");
            ttHelp.SetToolTip(btnSave, "boton para poder guardar los datos registrados en la aplicacion");
            //tt de ayuda para textbox
            ttHelp.SetToolTip(tcData, "Los cuadros de texto solo se deben llenar en una pestaña\nlas otras se actualizan en tiempo real");
            //llenado de tablas de descripcion
            for (i = 0; i < SDC.GetLength(0); i++)
            {
                dgvDesc1.Rows.Add(SDC[i, 0], SDC[i, 1], SDC[i, 2]);
                dgvDesc2.Rows.Add(SDC[i, 0], SDC[i, 1], SDC[i, 2]);
                dgvDesc3.Rows.Add(SDC[i, 0], SDC[i, 1], SDC[i, 2]);
                dgvDesc4.Rows.Add(SDC[i, 0], SDC[i, 1], SDC[i, 2]);
                dgvDesc5.Rows.Add(SDC[i, 0], SDC[i, 1], SDC[i, 2]);
            }
            //llenado y limitacion de tablas de preguntas
            for (i = 0; i < SDP1.GetLength(0); i++)
            {
                dgvPreg1.Rows.Add(SDP1[i, 0], SDP1[i, 1]);
                if (i == 0 || i == 4 || i == 8 || i == 12 || i == 16 || i == 22)
                {
                    dgvPreg1.Rows[i].ReadOnly = true;
                }
            }
            for (i = 0; i < SDP2.GetLength(0); i++)
            {
                dgvPreg2.Rows.Add(SDP2[i, 0], SDP2[i, 1]);
                if (i == 0 || i == 6 || i == 12 || i == 18)
                {
                    dgvPreg2.Rows[i].ReadOnly = true;
                }
            }
            for (i = 0; i < SDP3.GetLength(0); i++)
            {
                dgvPreg3.Rows.Add(SDP3[i, 0], SDP3[i, 1]);
                if (i == 0 || i == 4 || i == 8 || i == 12 || i == 16 || i == 22)
                {
                    dgvPreg3.Rows[i].ReadOnly = true;
                }
            }
            for (i = 0; i < SDP4.GetLength(0); i++)
            {
                dgvPreg4.Rows.Add(SDP4[i, 0], SDP4[i, 1]);
                if (i == 0 || i == 4 || i == 8 || i == 12 || i == 16 || i == 22)
                {
                    dgvPreg4.Rows[i].ReadOnly = true;
                }
            }
            for (i = 0; i < SDP5.GetLength(0); i++)
            {
                dgvPreg5.Rows.Add(SDP5[i, 0], SDP5[i, 1]);
                if (i == 0 || i == 5 || i == 10 || i == 15 || i == 19 || i == 23)
                {
                    dgvPreg5.Rows[i].ReadOnly = true;
                }
            }
            gbSave.Enabled = false;
        }
        //mouse move para calcular totales y total general
        private void btnCalcDat1_MouseMove(object sender, MouseEventArgs e)
        {
            a.Un = 0;
            a.Dos = 0;
            a.Tr = 0;
            a.Cu = 0;
            a.Cin = 0;

            uno = int.Parse(lbST11.Text.ToString()) + int.Parse(lbST12.Text.ToString()) + int.Parse(lbST13.Text.ToString()) + int.Parse(lbST14.Text.ToString()) + int.Parse(lbST15.Text.ToString());
            dos = int.Parse(lbST21.Text.ToString()) + int.Parse(lbST22.Text.ToString()) + int.Parse(lbST23.Text.ToString()) + int.Parse(lbST24.Text.ToString()) + int.Parse(lbST25.Text.ToString());
            tres = int.Parse(lbST31.Text.ToString()) + int.Parse(lbST32.Text.ToString()) + int.Parse(lbST33.Text.ToString()) + int.Parse(lbST34.Text.ToString()) + int.Parse(lbST35.Text.ToString());
            cuatro = int.Parse(lbST41.Text.ToString()) + int.Parse(lbST42.Text.ToString()) + int.Parse(lbST43.Text.ToString()) + int.Parse(lbST44.Text.ToString()) + int.Parse(lbST45.Text.ToString());
            cinco = int.Parse(lbST51.Text.ToString()) + int.Parse(lbST52.Text.ToString()) + int.Parse(lbST53.Text.ToString()) + int.Parse(lbST54.Text.ToString()) + int.Parse(lbST55.Text.ToString());

            a.U = uno;
            a.D = dos;
            a.T = tres;
            a.C = cuatro;
            a.Ci = cinco;

            a.totales();

            temp1 = string.Format("{0:0}", a.Subu);
            temp2 = string.Format("{0:0}", a.Subd);
            temp3 = string.Format("{0:0}", a.Subt);
            temp4 = string.Format("{0:0}", a.Subc);
            temp5 = string.Format("{0:0}", a.Subci);

            lbT11.Text = temp1;
            lbT12.Text = temp1;
            lbT13.Text = temp1;
            lbT14.Text = temp1;
            lbT15.Text = temp1;

            lbT21.Text = temp2;
            lbT22.Text = temp2;
            lbT23.Text = temp2;
            lbT24.Text = temp2;
            lbT25.Text = temp2;

            lbT31.Text = temp3;
            lbT32.Text = temp3;
            lbT33.Text = temp3;
            lbT34.Text = temp3;
            lbT35.Text = temp3;

            lbT41.Text = temp4;
            lbT42.Text = temp4;
            lbT43.Text = temp4;
            lbT44.Text = temp4;
            lbT45.Text = temp4;

            lbT51.Text = temp5;
            lbT52.Text = temp5;
            lbT53.Text = temp5;
            lbT54.Text = temp5;
            lbT55.Text = temp5;

            a.Un = double.Parse(lbT11.Text);
            a.Dos = double.Parse(lbT21.Text);
            a.Tr = double.Parse(lbT31.Text);
            a.Cu = double.Parse(lbT41.Text);
            a.Cin = double.Parse(lbT51.Text.ToString());

            a.totales();

            lbNC1.Text = String.Format("{0:0.0}", a.Total);
            lbNC2.Text = String.Format("{0:0.0}", a.Total);
            lbNC3.Text = String.Format("{0:0.0}", a.Total);
            lbNC4.Text = String.Format("{0:0.0}", a.Total);
            lbNC5.Text = String.Format("{0:0.0}", a.Total);
        }
        //activar groupbox
        private void tcData_MouseClick(object sender, MouseEventArgs e)
        {
            errorCom = "";
            //tabla 1
            for (i = 0; i < SDP1.GetLength(0); i++)
            {
                if (i != 0 && i != 4 && i != 8 && i != 12 && i != 16 && i != 22)
                {
                    if (dgvPreg1.Rows[i].Cells[2].Value != null)
                    {
                        pollo1 = true;
                    }
                    else
                    {
                        pollo1 = false;
                        errorCom = "llenar campos en la tabla de preguntas numero 1\n";
                        break;
                    }
                }
            }
            //tabla 2
            for (i = 0; i < SDP2.GetLength(0); i++)
            {
                if (i != 0 && i != 6 && i != 12 && i != 18)
                {
                    if (dgvPreg2.Rows[i].Cells[2].Value != null)
                    {
                        pollo2 = true;
                    }
                    else
                    {
                        pollo2 = false;
                        errorCom = errorCom + "llenar campos en la tabla de preguntas numero 2\n";
                        break;
                    }
                }
            }
            //tabla 3
            for (i = 0; i < SDP3.GetLength(0); i++)
            {
                if (i != 0 && i != 4 && i != 8 && i != 12 && i != 16 && i != 22)
                {
                    if (dgvPreg3.Rows[i].Cells[2].Value != null)
                    {
                        pollo3 = true;
                    }
                    else
                    {
                        pollo3 = false;
                        errorCom = errorCom + "llenar campos en la tabla de preguntas numero 3\n";
                        break;
                    }
                }
            }
            //tabla 4
            for (i = 0; i < SDP4.GetLength(0); i++)
            {
                if (i != 0 && i != 4 && i != 8 && i != 12 && i != 16 && i != 22)
                {
                    if (dgvPreg4.Rows[i].Cells[2].Value != null)
                    {
                        pollo4 = true;
                    }
                    else
                    {
                        pollo4 = false;
                        errorCom = errorCom + "llenar campos en la tabla de preguntas numero 4\n";
                        break;
                    }
                }
            }
            //tabla 5
            for (i = 0; i < SDP5.GetLength(0); i++)
            {
                if (i != 0 && i != 5 && i != 10 && i != 15 && i != 19 && i != 23)
                {
                    if (dgvPreg5.Rows[i].Cells[2].Value != null)
                    {
                        pollo5 = true;
                    }
                    else
                    {
                        pollo5 = false;
                        errorCom = errorCom + "llenar campos en la tabla de preguntas numero 5\n";
                        break;
                    }
                }
            }
            btnCalcDat1_Click(this, e);
            btnCalcDat2_Click(this, e);
            btnCalcDat3_Click(this, e);
            btnCalcDat4_Click(this, e);
            btnCalcDat5_Click(this, e);
            btnCalcDat1_MouseMove(this, e);
            epFail.Clear();
            if (textBox1.Text == String.Empty || textBox2.Text == String.Empty || textBox3.Text == String.Empty || textBox4.Text == String.Empty || textBox5.Text == String.Empty || textBox6.Text == String.Empty)
            {
                pollo1 = false;
                errorCom = errorCom + "llenar campos de texto en la aplicacion\n";
            }
            if (errorCom != string.Empty)
            {
                epFail.SetError(gbSave, "para poder usar este campo debe\n" + errorCom);
            }

            if (pollo1 == true && pollo2 == true && pollo3 == true && pollo4 == true && pollo5 == true)
            {
                epFail.Clear();
                gbSave.Enabled = true;
                rbDefault.Select();
                txtName.Enabled = false;
            }
        }
        //seleccionar radiobottoms
        private void rbDefault_CheckedChanged(object sender, EventArgs e)
        {
            txtName.Enabled = false;
            label22.Text = string.Empty;
        }
        private void rbpersonalized_CheckedChanged(object sender, EventArgs e)
        {
            txtName.Enabled = true;
            lbNombre.Enabled = false;
            label22.Text = string.Empty;
        }
        //btn guardar datos

        private void btnSave_Click(object sender, EventArgs e)
        {
            bool papas;
            int fritas;
            string name = null, dir = null;
            if (rbDefault.Checked)
            {
                name = lbNombre.Text.Trim();
                dir = string.Format(@"C:\Users\Public\Documents\" + name + ".docs");                
                Wrt = new StreamWriter(dir, false);
                //tabla 1
                Wrt.WriteLine(tp1.Text);
                Wrt.WriteLine("DATOS DE IDENTIFICACION");

                Wrt.WriteLine(label2.Text + " " + textBox6.Text + "\t\t\tCODIGO:" + textBox1.Text);
                Wrt.WriteLine(label3.Text + " " + textBox5.Text + "\t\t\tCODIGO:" + textBox2.Text);
                Wrt.WriteLine(label4.Text + " " + textBox4.Text + "\tCODIGO:" + textBox3.Text);
                Wrt.WriteLine("\n" + label5.Text + " " + dateTimePicker1.Value.ToString());

                Wrt.WriteLine("\n" + label9.Text + " " + label10.Text + "\n\n" + label11.Text + "\n");
                Wrt.WriteLine("Cualitativo\tDescripcion\t\tCualitativo");
                for (i = 0; i < SDC.GetLength(0); i++)
                {
                    Wrt.WriteLine(dgvDesc1.Rows[i].Cells[0].Value.ToString() + "\t\t" + dgvDesc1.Rows[i].Cells[1].Value.ToString() + "\t\t" + dgvDesc1.Rows[i].Cells[2].Value.ToString());
                }
                Wrt.WriteLine("IV\tCOMPONENTES TECNICOS ADMINISTRATIVOS \t\tSISTEMA DE CALIFICACION");
                for (i = 0; i < SDP1.GetLength(0); i++)
                {

                    if (i == 0)
                    {
                        Wrt.WriteLine("A. Prevision");
                    }
                    else if (i == 4)
                    {
                        Wrt.WriteLine("B. Planeacion");
                    }
                    else if (i == 8)
                    {
                        Wrt.WriteLine("C. Organizacion");
                    }
                    else if (i == 12)
                    {
                        Wrt.WriteLine("D. Integracion");
                    }
                    else if (i == 16)
                    {
                        Wrt.WriteLine("E. Dirreccion");
                    }
                    else if (i == 22)
                    {
                        Wrt.WriteLine("F. Control");
                    }
                    else if (i != 0 && i != 4 && i != 8 && i != 12 && i != 16 && i != 22)
                    {
                        Wrt.WriteLine(dgvPreg1.Rows[i].Cells[0].Value.ToString() + " " + dgvPreg1.Rows[i].Cells[1].Value.ToString() + "\t" + dgvPreg1.Rows[i].Cells[2].Value.ToString());
                    }
                }
                Wrt.WriteLine("\n" + label12.Text + " " + label13.Text + " " + label14.Text + " " + label15.Text + " " + label16.Text + " " + label17.Text);
                Wrt.WriteLine(label18.Text + "\t\t\t\t" + lbST11.Text + "\t\t" + lbST21.Text + "\t" + lbST31.Text + "\t" + lbST41.Text + "\t\t" + lbST51.Text);
                Wrt.WriteLine(label19.Text + "\t\t\t\t" + lbT11.Text + "\t\t" + lbT21.Text + "\t" + lbT31.Text + "\t" + lbT41.Text + "\t\t" + lbT51.Text);
                Wrt.WriteLine(label20.Text + "\t" + lbNC1.Text);

                //tabla 2
                Wrt.WriteLine("\n");
                Wrt.WriteLine(tp2.Text);
                Wrt.WriteLine("DATOS DE IDENTIFICACION");

                Wrt.WriteLine(label2.Text + " " + textBox6.Text + "\t\t\tCODIGO:" + textBox1.Text);
                Wrt.WriteLine(label3.Text + " " + textBox7.Text + "\t\t\tCODIGO:" + textBox2.Text);
                Wrt.WriteLine(label4.Text + " " + textBox8.Text + "\tCODIGO:" + textBox3.Text);
                Wrt.WriteLine("\n" + label5.Text + " " + dateTimePicker2.Value.ToString());

                Wrt.WriteLine("\n" + label9.Text + " " + label10.Text + "\n\n" + label11.Text + "\n");
                Wrt.WriteLine("Cualitativo\tDescripcion\t\tCualitativo");
                for (i = 0; i < SDC.GetLength(0); i++)
                {
                    Wrt.WriteLine(dgvDesc1.Rows[i].Cells[0].Value.ToString() + "\t\t" + dgvDesc1.Rows[i].Cells[1].Value.ToString() + "\t\t" + dgvDesc1.Rows[i].Cells[2].Value.ToString());
                }
                Wrt.WriteLine("IV\tCOMPONENTES TECNICOS ADMINISTRATIVOS \t\tSISTEMA DE CALIFICACION");
                for (i = 0; i < SDP2.GetLength(0); i++)
                {

                    if (i == 0)
                    {
                        Wrt.WriteLine("A. Formulacion de Planos");
                    }
                    else if (i == 6)
                    {
                        Wrt.WriteLine("B. Control Administrativo Ejercido");
                    }
                    else if (i == 12)
                    {
                        Wrt.WriteLine("C. Manera de Dirigir");
                    }
                    else if (i == 18)
                    {
                        Wrt.WriteLine("D. Comunicacion Efectiva");
                    }
                    else if (i != 0 && i != 6 && i != 12 && i != 18)
                    {
                        Wrt.WriteLine(dgvPreg2.Rows[i].Cells[0].Value.ToString() + " " + dgvPreg2.Rows[i].Cells[1].Value.ToString() + "\t" + dgvPreg2.Rows[i].Cells[2].Value.ToString());
                    }
                }
                Wrt.WriteLine("\n" + label12.Text + " " + label13.Text + " " + label14.Text + " " + label15.Text + " " + label16.Text + " " + label17.Text);
                Wrt.WriteLine(label18.Text + "\t\t\t\t" + lbST12.Text + "\t\t" + lbST22.Text + "\t" + lbST32.Text + "\t" + lbST42.Text + "\t\t" + lbST52.Text);
                Wrt.WriteLine(label19.Text + "\t\t\t\t" + lbT11.Text + "\t\t" + lbT21.Text + "\t" + lbT31.Text + "\t" + lbT41.Text + "\t\t" + lbT51.Text);
                Wrt.WriteLine(label20.Text + "\t" + lbNC1.Text);

                //tabla 3
                Wrt.WriteLine("\n");
                Wrt.WriteLine(tp3.Text);
                Wrt.WriteLine("DATOS DE IDENTIFICACION");

                Wrt.WriteLine(label2.Text + " " + textBox6.Text + "\t\t\tCODIGO:" + textBox1.Text);
                Wrt.WriteLine(label3.Text + " " + textBox7.Text + "\t\t\tCODIGO:" + textBox2.Text);
                Wrt.WriteLine(label4.Text + " " + textBox8.Text + "\tCODIGO:" + textBox3.Text);
                Wrt.WriteLine("\n" + label5.Text + " " + dateTimePicker3.Value.ToString());

                Wrt.WriteLine("\n" + label9.Text + " " + label10.Text + "\n\n" + label11.Text + "\n");
                Wrt.WriteLine("Cualitativo\tDescripcion\t\tCualitativo");
                for (i = 0; i < SDC.GetLength(0); i++)
                {
                    Wrt.WriteLine(dgvDesc1.Rows[i].Cells[0].Value.ToString() + "\t\t" + dgvDesc1.Rows[i].Cells[1].Value.ToString() + "\t\t" + dgvDesc1.Rows[i].Cells[2].Value.ToString());
                }
                Wrt.WriteLine("IV\tCOMPONENTES TECNICOS ADMINISTRATIVOS \t\tSISTEMA DE CALIFICACION");
                for (i = 0; i < SDP3.GetLength(0); i++)
                {

                    if (i == 0)
                    {
                        Wrt.WriteLine("A. Motivacion");
                    }
                    else if (i == 4)
                    {
                        Wrt.WriteLine("B. Liderazgo");
                    }
                    else if (i == 8)
                    {
                        Wrt.WriteLine("C. Comunicacion");
                    }
                    else if (i == 12)
                    {
                        Wrt.WriteLine("D. Desarrollo de Equipo");
                    }
                    else if (i == 16)
                    {
                        Wrt.WriteLine("E. Gestion de Cambio");
                    }
                    else if (i == 22)
                    {
                        Wrt.WriteLine("F. Vision Estrategica");
                    }
                    else if (i != 0 && i != 4 && i != 8 && i != 12 && i != 16 && i != 22)
                    {
                        Wrt.WriteLine(dgvPreg3.Rows[i].Cells[0].Value.ToString() + " " + dgvPreg3.Rows[i].Cells[1].Value.ToString() + "\t" + dgvPreg3.Rows[i].Cells[2].Value.ToString());
                    }
                }
                Wrt.WriteLine("\n" + label12.Text + " " + label13.Text + " " + label14.Text + " " + label15.Text + " " + label16.Text + " " + label17.Text);
                Wrt.WriteLine(label18.Text + "\t\t\t\t" + lbST13.Text + "\t\t" + lbST23.Text + "\t" + lbST33.Text + "\t" + lbST43.Text + "\t\t" + lbST53.Text);
                Wrt.WriteLine(label19.Text + "\t\t\t\t" + lbT11.Text + "\t\t" + lbT21.Text + "\t" + lbT31.Text + "\t" + lbT41.Text + "\t\t" + lbT51.Text);
                Wrt.WriteLine(label20.Text + "\t" + lbNC1.Text);

                //tabla 4
                Wrt.WriteLine("\n");
                Wrt.WriteLine(tp4.Text);
                Wrt.WriteLine("DATOS DE IDENTIFICACION");

                Wrt.WriteLine(label2.Text + " " + textBox6.Text + "\t\t\tCODIGO:" + textBox1.Text);
                Wrt.WriteLine(label3.Text + " " + textBox7.Text + "\t\t\tCODIGO:" + textBox2.Text);
                Wrt.WriteLine(label4.Text + " " + textBox8.Text + "\tCODIGO:" + textBox3.Text);
                Wrt.WriteLine("\n" + label5.Text + " " + dateTimePicker4.Value.ToString());

                Wrt.WriteLine("\n" + label9.Text + " " + label10.Text + "\n\n" + label11.Text + "\n");
                Wrt.WriteLine("Cualitativo\tDescripcion\t\tCualitativo");
                for (i = 0; i < SDC.GetLength(0); i++)
                {
                    Wrt.WriteLine(dgvDesc1.Rows[i].Cells[0].Value.ToString() + "\t\t" + dgvDesc1.Rows[i].Cells[1].Value.ToString() + "\t\t" + dgvDesc1.Rows[i].Cells[2].Value.ToString());
                }
                Wrt.WriteLine("IV\tCOMPONENTES TECNICOS ADMINISTRATIVOS \t\tSISTEMA DE CALIFICACION");
                for (i = 0; i < SDP4.GetLength(0); i++)
                {

                    if (i == 0)
                    {
                        Wrt.WriteLine("A. Democracia");
                    }
                    else if (i == 4)
                    {
                        Wrt.WriteLine("B. Participacion");
                    }
                    else if (i == 8)
                    {
                        Wrt.WriteLine("C. Situacional");
                    }
                    else if (i == 12)
                    {
                        Wrt.WriteLine("D. Maternidad o Paternida");
                    }
                    else if (i == 16)
                    {
                        Wrt.WriteLine("E. Contingencial");
                    }
                    else if (i == 22)
                    {
                        Wrt.WriteLine("F. Transformacional");
                    }
                    else if (i != 0 && i != 4 && i != 8 && i != 12 && i != 16 && i != 22)
                    {
                        Wrt.WriteLine(dgvPreg4.Rows[i].Cells[0].Value.ToString() + " " + dgvPreg4.Rows[i].Cells[1].Value.ToString() + "\t" + dgvPreg4.Rows[i].Cells[2].Value.ToString());
                    }
                }
                Wrt.WriteLine("\n" + label12.Text + " " + label13.Text + " " + label14.Text + " " + label15.Text + " " + label16.Text + " " + label17.Text);
                Wrt.WriteLine(label18.Text + "\t\t\t\t" + lbST14.Text + "\t\t" + lbST24.Text + "\t" + lbST34.Text + "\t" + lbST44.Text + "\t\t" + lbST54.Text);
                Wrt.WriteLine(label19.Text + "\t\t\t\t" + lbT11.Text + "\t\t" + lbT21.Text + "\t" + lbT31.Text + "\t" + lbT41.Text + "\t\t" + lbT51.Text);
                Wrt.WriteLine(label20.Text + "\t" + lbNC1.Text);

                //tabla 5
                Wrt.WriteLine("\n");
                Wrt.WriteLine(tp5.Text);
                Wrt.WriteLine("DATOS DE IDENTIFICACION");

                Wrt.WriteLine(label2.Text + " " + textBox6.Text + "\t\t\tCODIGO:" + textBox1.Text);
                Wrt.WriteLine(label3.Text + " " + textBox7.Text + "\t\t\tCODIGO:" + textBox2.Text);
                Wrt.WriteLine(label4.Text + " " + textBox8.Text + "\tCODIGO:" + textBox3.Text);
                Wrt.WriteLine("\n" + label5.Text + " " + dateTimePicker5.Value.ToString());

                Wrt.WriteLine("\n" + label9.Text + " " + label10.Text + "\n\n" + label11.Text + "\n");
                Wrt.WriteLine("Cualitativo\tDescripcion\t\tCualitativo");
                for (i = 0; i < SDC.GetLength(0); i++)
                {
                    Wrt.WriteLine(dgvDesc1.Rows[i].Cells[0].Value.ToString() + "\t\t" + dgvDesc1.Rows[i].Cells[1].Value.ToString() + "\t\t" + dgvDesc1.Rows[i].Cells[2].Value.ToString());
                }
                Wrt.WriteLine("IV\tCOMPONENTES TECNICOS ADMINISTRATIVOS \t\tSISTEMA DE CALIFICACION");
                for (i = 0; i < SDP5.GetLength(0); i++)
                {

                    if (i == 0)
                    {
                        Wrt.WriteLine("A. Red de apoyo");
                    }
                    else if (i == 5)
                    {
                        Wrt.WriteLine("B. Manera de Comunicarse");
                    }
                    else if (i == 10)
                    {
                        Wrt.WriteLine("C. Forma de Delegar");
                    }
                    else if (i == 15)
                    {
                        Wrt.WriteLine("D. Manera de Evaluacion");
                    }
                    else if (i == 19)
                    {
                        Wrt.WriteLine("E. Manera de Escucha Activa");
                    }
                    else if (i == 23)
                    {
                        Wrt.WriteLine("F. Generacion de Entorno Inclusivo");
                    }
                    else if (i != 0 && i != 5 && i != 10 && i != 15 && i != 19 && i != 23)
                    {
                        Wrt.WriteLine(dgvPreg5.Rows[i].Cells[0].Value.ToString() + " " + dgvPreg5.Rows[i].Cells[1].Value.ToString() + "\t" + dgvPreg5.Rows[i].Cells[2].Value.ToString());
                    }
                }
                Wrt.WriteLine("\n" + label12.Text + " " + label13.Text + " " + label14.Text + " " + label15.Text + " " + label16.Text + " " + label17.Text);
                Wrt.WriteLine(label18.Text + "\t" + lbST13.Text + "\t" + lbST23.Text + "\t" + lbST33.Text + "\t" + lbST43.Text + "\t" + lbST53.Text);
                Wrt.WriteLine(label19.Text + "\t\t\t\t" + lbT11.Text + "\t\t" + lbT21.Text + "\t" + lbT31.Text + "\t" + lbT41.Text + "\t\t" + lbT51.Text);
                Wrt.WriteLine(label20.Text + "\t" + lbNC1.Text);

                Wrt.Close();

                papas = int.TryParse(Microsoft.VisualBasic.Interaction.InputBox("Desea abrir el archivo que se acaba de crear\n1. SI\n2. NO"), out fritas);
                if (papas == true)
                {
                    if (fritas == 1)
                    {
                        Process.Start(dir);
                    }
                    else if (fritas != 1)
                    {
                        Microsoft.VisualBasic.Interaction.MsgBox("EL ARCHIVO NO SE ABRIRA\tSE LE MOSTRARA A CONTINUACION LA RUTA DE GUARDADO DEL ARCHIVO (para abrirlo puede hacerlo como un documento de word)");
                    }
                }
                else if (papas == false)
                {
                    Microsoft.VisualBasic.Interaction.MsgBox("NO PUDIMOS PROCESAR SU SOLICITUD\nSI DESEA ABRIR EL ARCHIVO DEJAREMOS LA DIRECCION EN LA PANTALLA DE GUARDADO");
                }
                label22.Text = "SU DOCUMENTO SE GUARDO EXITOSAMENTE EN LA DIRECCION: " + string.Format(@"C:\Users\Public\Documents") + "\nBAJO EL NOMBRE: " + name;
            }
            if (rbpersonalized.Checked && txtName.Text != String.Empty)
            {
                name = txtName.Text.Trim();
                dir = string.Format(@"C:\Users\Public\Documents\" + name + ".docs");

                Wrt = new StreamWriter(dir, false);
                //tabla 1
                Wrt.WriteLine("\n");
                Wrt.WriteLine(tp1.Text);
                Wrt.WriteLine("DATOS DE IDENTIFICACION");

                Wrt.WriteLine(label2.Text + " " + textBox6.Text + "\t\t\tCODIGO:" + textBox1.Text);
                Wrt.WriteLine(label3.Text + " " + textBox7.Text + "\t\t\tCODIGO:" + textBox2.Text);
                Wrt.WriteLine(label4.Text + " " + textBox8.Text + "\tCODIGO:" + textBox3.Text);
                Wrt.WriteLine("\n" + label5.Text + " " + dateTimePicker1.Value.ToString());

                Wrt.WriteLine("\n" + label9.Text + " " + label10.Text + "\n\n" + label11.Text + "\n");
                Wrt.WriteLine("Cualitativo\tDescripcion\t\tCualitativo");
                for (i = 0; i < SDC.GetLength(0); i++)
                {
                    Wrt.WriteLine(dgvDesc1.Rows[i].Cells[0].Value.ToString() + "\t\t" + dgvDesc1.Rows[i].Cells[1].Value.ToString() + "\t\t" + dgvDesc1.Rows[i].Cells[2].Value.ToString());
                }
                Wrt.WriteLine("IV\tCOMPONENTES TECNICOS ADMINISTRATIVOS \t\tSISTEMA DE CALIFICACION");
                for (i = 0; i < SDP1.GetLength(0); i++)
                {

                    if (i == 0)
                    {
                        Wrt.WriteLine("A. Prevision");
                    }
                    else if (i == 4)
                    {
                        Wrt.WriteLine("B. Planeacion");
                    }
                    else if (i == 8)
                    {
                        Wrt.WriteLine("C. Organizacion");
                    }
                    else if (i == 12)
                    {
                        Wrt.WriteLine("D. Integracion");
                    }
                    else if (i == 16)
                    {
                        Wrt.WriteLine("E. Dirreccion");
                    }
                    else if (i == 22)
                    {
                        Wrt.WriteLine("F. Control");
                    }
                    else if (i != 0 && i != 4 && i != 8 && i != 12 && i != 16 && i != 22)
                    {
                        Wrt.WriteLine(dgvPreg1.Rows[i].Cells[0].Value.ToString() + " " + dgvPreg1.Rows[i].Cells[1].Value.ToString() + "\t" + dgvPreg1.Rows[i].Cells[2].Value.ToString());
                    }
                }
                Wrt.WriteLine("\n" + label12.Text + " " + label13.Text + " " + label14.Text + " " + label15.Text + " " + label16.Text + " " + label17.Text);
                Wrt.WriteLine(label18.Text + "\t\t\t\t" + lbST11.Text + "\t\t" + lbST21.Text + "\t" + lbST31.Text + "\t" + lbST41.Text + "\t\t" + lbST51.Text);
                Wrt.WriteLine(label19.Text + "\t\t\t\t" + lbT11.Text + "\t\t" + lbT21.Text + "\t" + lbT31.Text + "\t" + lbT41.Text + "\t\t" + lbT51.Text);
                Wrt.WriteLine(label20.Text + "\t" + lbNC1.Text);

                //tabla 2
                Wrt.WriteLine("\n");
                Wrt.WriteLine(tp2.Text);
                Wrt.WriteLine("DATOS DE IDENTIFICACION");

                Wrt.WriteLine(label2.Text + " " + textBox6.Text + "\t\t\tCODIGO:" + textBox1.Text);
                Wrt.WriteLine(label3.Text + " " + textBox7.Text + "\t\t\tCODIGO:" + textBox2.Text);
                Wrt.WriteLine(label4.Text + " " + textBox8.Text + "\tCODIGO:" + textBox3.Text);
                Wrt.WriteLine("\n" + label5.Text + " " + dateTimePicker2.Value.ToString());

                Wrt.WriteLine("\n" + label9.Text + " " + label10.Text + "\n\n" + label11.Text + "\n");
                Wrt.WriteLine("Cualitativo\tDescripcion\t\tCualitativo");
                for (i = 0; i < SDC.GetLength(0); i++)
                {
                    Wrt.WriteLine(dgvDesc1.Rows[i].Cells[0].Value.ToString() + "\t\t" + dgvDesc1.Rows[i].Cells[1].Value.ToString() + "\t\t" + dgvDesc1.Rows[i].Cells[2].Value.ToString());
                }
                Wrt.WriteLine("IV\tCOMPONENTES TECNICOS ADMINISTRATIVOS \t\tSISTEMA DE CALIFICACION");
                for (i = 0; i < SDP2.GetLength(0); i++)
                {

                    if (i == 0)
                    {
                        Wrt.WriteLine("A. Formulacion de Planos");
                    }
                    else if (i == 6)
                    {
                        Wrt.WriteLine("B. Control Administrativo Ejercido");
                    }
                    else if (i == 12)
                    {
                        Wrt.WriteLine("C. Manera de Dirigir");
                    }
                    else if (i == 18)
                    {
                        Wrt.WriteLine("D. Comunicacion Efectiva");
                    }
                    else if (i != 0 && i != 6 && i != 12 && i != 18)
                    {
                        Wrt.WriteLine(dgvPreg2.Rows[i].Cells[0].Value.ToString() + " " + dgvPreg2.Rows[i].Cells[1].Value.ToString() + "\t" + dgvPreg2.Rows[i].Cells[2].Value.ToString());
                    }
                }
                Wrt.WriteLine("\n" + label12.Text + " " + label13.Text + " " + label14.Text + " " + label15.Text + " " + label16.Text + " " + label17.Text);
                Wrt.WriteLine(label18.Text + "\t\t\t\t" + lbST12.Text + "\t\t" + lbST22.Text + "\t" + lbST32.Text + "\t" + lbST42.Text + "\t\t" + lbST52.Text);
                Wrt.WriteLine(label19.Text + "\t\t\t\t" + lbT11.Text + "\t\t" + lbT21.Text + "\t" + lbT31.Text + "\t" + lbT41.Text + "\t\t" + lbT51.Text);
                Wrt.WriteLine(label20.Text + "\t" + lbNC1.Text);

                //tabla 3
                Wrt.WriteLine("\n");
                Wrt.WriteLine(tp3.Text);
                Wrt.WriteLine("DATOS DE IDENTIFICACION");

                Wrt.WriteLine(label2.Text + " " + textBox6.Text + "\t\t\tCODIGO:" + textBox1.Text);
                Wrt.WriteLine(label3.Text + " " + textBox7.Text + "\t\t\tCODIGO:" + textBox2.Text);
                Wrt.WriteLine(label4.Text + " " + textBox8.Text + "\tCODIGO:" + textBox3.Text);
                Wrt.WriteLine("\n" + label5.Text + " " + dateTimePicker3.Value.ToString());

                Wrt.WriteLine("\n" + label9.Text + " " + label10.Text + "\n\n" + label11.Text + "\n");
                Wrt.WriteLine("Cualitativo\tDescripcion\t\tCualitativo");
                for (i = 0; i < SDC.GetLength(0); i++)
                {
                    Wrt.WriteLine(dgvDesc1.Rows[i].Cells[0].Value.ToString() + "\t\t" + dgvDesc1.Rows[i].Cells[1].Value.ToString() + "\t\t" + dgvDesc1.Rows[i].Cells[2].Value.ToString());
                }
                Wrt.WriteLine("IV\tCOMPONENTES TECNICOS ADMINISTRATIVOS \t\tSISTEMA DE CALIFICACION");
                for (i = 0; i < SDP3.GetLength(0); i++)
                {

                    if (i == 0)
                    {
                        Wrt.WriteLine("A. Motivacion");
                    }
                    else if (i == 4)
                    {
                        Wrt.WriteLine("B. Liderazgo");
                    }
                    else if (i == 8)
                    {
                        Wrt.WriteLine("C. Comunicacion");
                    }
                    else if (i == 12)
                    {
                        Wrt.WriteLine("D. Desarrollo de Equipo");
                    }
                    else if (i == 16)
                    {
                        Wrt.WriteLine("E. Gestion de Cambio");
                    }
                    else if (i == 22)
                    {
                        Wrt.WriteLine("F. Vision Estrategica");
                    }
                    else if (i != 0 && i != 4 && i != 8 && i != 12 && i != 16 && i != 22)
                    {
                        Wrt.WriteLine(dgvPreg3.Rows[i].Cells[0].Value.ToString() + " " + dgvPreg3.Rows[i].Cells[1].Value.ToString() + "\t" + dgvPreg3.Rows[i].Cells[2].Value.ToString());
                    }
                }
                Wrt.WriteLine("\n" + label12.Text + " " + label13.Text + " " + label14.Text + " " + label15.Text + " " + label16.Text + " " + label17.Text);
                Wrt.WriteLine(label18.Text + "\t\t\t\t" + lbST13.Text + "\t\t" + lbST23.Text + "\t" + lbST33.Text + "\t" + lbST43.Text + "\t\t" + lbST53.Text);
                Wrt.WriteLine(label19.Text + "\t\t\t\t" + lbT11.Text + "\t\t" + lbT21.Text + "\t" + lbT31.Text + "\t" + lbT41.Text + "\t\t" + lbT51.Text);
                Wrt.WriteLine(label20.Text + "\t" + lbNC1.Text);

                //tabla 4
                Wrt.WriteLine("\n");
                Wrt.WriteLine(tp4.Text);
                Wrt.WriteLine("DATOS DE IDENTIFICACION");

                Wrt.WriteLine(label2.Text + " " + textBox6.Text + "\t\t\tCODIGO:" + textBox1.Text);
                Wrt.WriteLine(label3.Text + " " + textBox7.Text + "\t\t\tCODIGO:" + textBox2.Text);
                Wrt.WriteLine(label4.Text + " " + textBox8.Text + "\tCODIGO:" + textBox3.Text);
                Wrt.WriteLine("\n" + label5.Text + " " + dateTimePicker4.Value.ToString());

                Wrt.WriteLine("\n" + label9.Text + " " + label10.Text + "\n\n" + label11.Text + "\n");
                Wrt.WriteLine("Cualitativo\tDescripcion\t\tCualitativo");
                for (i = 0; i < SDC.GetLength(0); i++)
                {
                    Wrt.WriteLine(dgvDesc1.Rows[i].Cells[0].Value.ToString() + "\t\t" + dgvDesc1.Rows[i].Cells[1].Value.ToString() + "\t\t" + dgvDesc1.Rows[i].Cells[2].Value.ToString());
                }
                Wrt.WriteLine("IV\tCOMPONENTES TECNICOS ADMINISTRATIVOS \t\tSISTEMA DE CALIFICACION");
                for (i = 0; i < SDP4.GetLength(0); i++)
                {

                    if (i == 0)
                    {
                        Wrt.WriteLine("A. Democracia");
                    }
                    else if (i == 4)
                    {
                        Wrt.WriteLine("B. Participacion");
                    }
                    else if (i == 8)
                    {
                        Wrt.WriteLine("C. Situacional");
                    }
                    else if (i == 12)
                    {
                        Wrt.WriteLine("D. Maternidad o Paternida");
                    }
                    else if (i == 16)
                    {
                        Wrt.WriteLine("E. Contingencial");
                    }
                    else if (i == 22)
                    {
                        Wrt.WriteLine("F. Transformacional");
                    }
                    else if (i != 0 && i != 4 && i != 8 && i != 12 && i != 16 && i != 22)
                    {
                        Wrt.WriteLine(dgvPreg4.Rows[i].Cells[0].Value.ToString() + " " + dgvPreg4.Rows[i].Cells[1].Value.ToString() + "\t" + dgvPreg4.Rows[i].Cells[2].Value.ToString());
                    }
                }
                Wrt.WriteLine("\n" + label12.Text + " " + label13.Text + " " + label14.Text + " " + label15.Text + " " + label16.Text + " " + label17.Text);
                Wrt.WriteLine(label18.Text + "\t\t\t\t" + lbST14.Text + "\t\t" + lbST24.Text + "\t" + lbST34.Text + "\t" + lbST44.Text + "\t\t" + lbST54.Text);
                Wrt.WriteLine(label19.Text + "\t\t\t\t" + lbT11.Text + "\t\t" + lbT21.Text + "\t" + lbT31.Text + "\t" + lbT41.Text + "\t\t" + lbT51.Text);
                Wrt.WriteLine(label20.Text + "\t" + lbNC1.Text);

                //tabla 5
                Wrt.WriteLine("\n");
                Wrt.WriteLine(tp5.Text);
                Wrt.WriteLine("DATOS DE IDENTIFICACION");

                Wrt.WriteLine(label2.Text + " " + textBox6.Text + "\t\t\tCODIGO:" + textBox1.Text);
                Wrt.WriteLine(label3.Text + " " + textBox7.Text + "\t\t\tCODIGO:" + textBox2.Text);
                Wrt.WriteLine(label4.Text + " " + textBox8.Text + "\tCODIGO:" + textBox3.Text);
                Wrt.WriteLine("\n" + label5.Text + " " + dateTimePicker5.Value.ToString());

                Wrt.WriteLine("\n" + label9.Text + " " + label10.Text + "\n\n" + label11.Text + "\n");
                Wrt.WriteLine("Cualitativo\tDescripcion\t\tCualitativo");
                for (i = 0; i < SDC.GetLength(0); i++)
                {
                    Wrt.WriteLine(dgvDesc1.Rows[i].Cells[0].Value.ToString() + "\t\t" + dgvDesc1.Rows[i].Cells[1].Value.ToString() + "\t\t" + dgvDesc1.Rows[i].Cells[2].Value.ToString());
                }
                Wrt.WriteLine("IV\tCOMPONENTES TECNICOS ADMINISTRATIVOS \t\tSISTEMA DE CALIFICACION");
                for (i = 0; i < SDP5.GetLength(0); i++)
                {

                    if (i == 0)
                    {
                        Wrt.WriteLine("A. Red de apoyo");
                    }
                    else if (i == 5)
                    {
                        Wrt.WriteLine("B. Manera de Comunicarse");
                    }
                    else if (i == 10)
                    {
                        Wrt.WriteLine("C. Forma de Delegar");
                    }
                    else if (i == 15)
                    {
                        Wrt.WriteLine("D. Manera de Evaluacion");
                    }
                    else if (i == 19)
                    {
                        Wrt.WriteLine("E. Manera de Escucha Activa");
                    }
                    else if (i == 23)
                    {
                        Wrt.WriteLine("F. Generacion de Entorno Inclusivo");
                    }
                    else if (i != 0 && i != 5 && i != 10 && i != 15 && i != 19 && i != 23)
                    {
                        Wrt.WriteLine(dgvPreg5.Rows[i].Cells[0].Value.ToString() + " " + dgvPreg5.Rows[i].Cells[1].Value.ToString() + "\t" + dgvPreg5.Rows[i].Cells[2].Value.ToString());
                    }
                }
                Wrt.WriteLine("\n" + label12.Text + " " + label13.Text + " " + label14.Text + " " + label15.Text + " " + label16.Text + " " + label17.Text);
                Wrt.WriteLine(label18.Text + "\t" + lbST13.Text + "\t" + lbST23.Text + "\t" + lbST33.Text + "\t" + lbST43.Text + "\t" + lbST53.Text);
                Wrt.WriteLine(label19.Text + "\t\t\t\t" + lbT11.Text + "\t\t" + lbT21.Text + "\t" + lbT31.Text + "\t" + lbT41.Text + "\t\t" + lbT51.Text);
                Wrt.WriteLine(label20.Text + "\t" + lbNC1.Text);

                Wrt.Close();

                papas = int.TryParse(Microsoft.VisualBasic.Interaction.InputBox("Desea abrir el archivo que se acaba de crear\n1. SI\n2. NO"), out fritas);
                if (papas == true)
                {
                    if (fritas == 1)
                    {
                        Process.Start(dir);
                    }
                    else if (fritas != 1)
                    {
                        Microsoft.VisualBasic.Interaction.MsgBox("EL ARCHIVO NO SE ABRIRA\tSE LE MOSTRARA A CONTINUACION LA RUTA DE GUARDADO DEL ARCHIVO (para abrirlo puede hacerlo como un documento de word)");
                    }
                }
                else if (papas == false)
                {
                    Microsoft.VisualBasic.Interaction.MsgBox("NO PUDIMOS PROCESAR SU SOLICITUD\nSI DESEA ABRIR EL ARCHIVO DEJAREMOS LA DIRECCION EN LA PANTALLA DE GUARDADO");
                }
                label22.Text = "SU DOCUMENTO SE GUARDO EXITOSAMENTE EN LA DIRECCION: " + string.Format(@"C:\Users\Public\Documents")+"\nBAJO EL NOMBRE: " + name;
            }
            else if (rbpersonalized.Checked && txtName.Text == String.Empty)
            {
                epFail.SetError(txtName, "Debe llenar este campo antes de poder continuar");
                Microsoft.VisualBasic.Interaction.MsgBox("DEBIDO A QUE FALTA EL NOMBRE DEL DOCUMENTO NO SE PUEDE CREAR EL ARCHIVO");
            }
        }
        //btn para llenado de pruebas
        private void btnRellenar_Click(object sender, EventArgs e)
        {
            for (i = 0; i < SDP1.GetLength(0); i++)
            {
                if (i != 0 && i != 4 && i != 8 && i != 12 && i != 16 && i != 22)
                    dgvPreg1.Rows[i].Cells[2].Value = "Excelente";
            }
            for (i = 0; i < SDP2.GetLength(0); i++)
            {
                if (i != 0 && i != 6 && i != 12 && i != 18)
                {
                    dgvPreg2.Rows[i].Cells[2].Value = "Muy Bueno";
                }
            }
            for (i = 0; i < SDP3.GetLength(0); i++)
            {
                if (i != 0 && i != 4 && i != 8 && i != 12 && i != 16 && i != 22)
                {
                    dgvPreg3.Rows[i].Cells[2].Value = "Bueno";
                }
            }
            for (i = 0; i < SDP4.GetLength(0); i++)
            {
                if (i != 0 && i != 4 && i != 8 && i != 12 && i != 16 && i != 22)
                {
                    dgvPreg4.Rows[i].Cells[2].Value = "Regular";
                }
            }
            for (i = 0; i < SDP5.GetLength(0); i++)
            {
                if (i != 0 && i != 5 && i != 10 && i != 15 && i != 19 && i != 23)
                {
                    dgvPreg5.Rows[i].Cells[2].Value = "Deficiente";
                }
            }
        }
        //btn's para calcular subtotales
        private void btnCalcDat1_Click(object sender, EventArgs e)
        {
            error = "";
            lbNC1.Focus();

            a.U = 0;
            a.D = 0;
            a.T = 0;
            a.C = 0;
            a.Ci = 0;

            for (i = 0; i < SDP1.GetLength(0); i++)
            {
                if (i != 0 && i != 4 && i != 8 && i != 12 && i != 16 && i != 22)
                {
                    if (i != 0 && dgvPreg1.Rows[i].Cells[2].Value != null)
                    {
                        if (dgvPreg1.Rows[i].Cells[2].Value.ToString() == "Deficiente")
                        {
                            a.U++;
                        }
                        else if (dgvPreg1.Rows[i].Cells[2].Value.ToString() == "Regular")
                        {
                            a.D++;
                        }
                        else if (dgvPreg1.Rows[i].Cells[2].Value.ToString() == "Bueno")
                        {
                            a.T++;
                        }
                        else if (dgvPreg1.Rows[i].Cells[2].Value.ToString() == "Muy Bueno")
                        {
                            a.C++;
                        }
                        else if (dgvPreg1.Rows[i].Cells[2].Value.ToString() == "Excelente")
                        {
                            a.Ci++;
                        }
                        
                        a.totales();

                        lbST11.Text = "" + a.U;
                        lbST21.Text = "" + a.D;
                        lbST31.Text = "" + a.T;
                        lbST41.Text = "" + a.C;
                        lbST51.Text = "" + a.Ci;
                    }
                    else if (i != 0 && i != 4 && i != 8 && i != 12 && i != 16 && i != 22 && dgvPreg1.Rows[i].Cells[2].Value == null)
                    {
                        error = error + "\n" + dgvPreg1.Rows[i].Cells[0].Value;                       
                    }
                }
            }

            if (error != string.Empty)
            {
                epFail.SetError(this.dgvPreg1, "todavia se debe llenar la pregunta" + error);
            }
            else if (error == string.Empty && textBox6.Text != string.Empty)
            {
                epFail.Clear();
            }
        }
        private void btnCalcDat2_Click(object sender, EventArgs e)
        {
            error = "";
            lbNC2.Focus();

            a.U = 0;
            a.D = 0;
            a.T = 0;
            a.C = 0;
            a.Ci = 0;

            for (i = 0; i < SDP2.GetLength(0); i++)
            {

                if (i != 0 && i != 6 && i != 12 && i != 18)
                {
                    if (i != 0 && dgvPreg2.Rows[i].Cells[2].Value != null)
                    {
                        if (dgvPreg2.Rows[i].Cells[2].Value.ToString() == "Deficiente")
                        {
                            a.U++;
                        }
                        else if (dgvPreg2.Rows[i].Cells[2].Value.ToString() == "Regular")
                        {
                            a.D++;
                        }
                        else if (dgvPreg2.Rows[i].Cells[2].Value.ToString() == "Bueno")
                        {
                            a.T++;
                        }
                        else if (dgvPreg2.Rows[i].Cells[2].Value.ToString() == "Muy Bueno")
                        {
                            a.C++;
                        }
                        else if (dgvPreg2.Rows[i].Cells[2].Value.ToString() == "Excelente")
                        {
                            a.Ci++;
                        }

                        a.totales();

                        lbST12.Text = "" + a.U;
                        lbST22.Text = "" + a.D;
                        lbST32.Text = "" + a.T;
                        lbST42.Text = "" + a.C;
                        lbST52.Text = "" + a.Ci;
                    }
                    else if (i != 0 && i != 6 && i != 12 && i != 18 && dgvPreg2.Rows[i].Cells[2].Value == null)
                    {
                        error = error + "\n" + dgvPreg2.Rows[i].Cells[0].Value;
                    }
                }
            }

            if (error != string.Empty)
            {
                epFail.SetError(this.dgvPreg2, "todavia se debe llenar la pregunta" + error);
            }
            else if (error == string.Empty)
            {
                epFail.Clear();
            }
        }
        private void btnCalcDat3_Click(object sender, EventArgs e)
        {
            error = "";
            lbNC3.Focus();

            a.U = 0;
            a.D = 0;
            a.T = 0;
            a.C = 0;
            a.Ci = 0;

            for (i = 0; i < SDP3.GetLength(0); i++)
            {

                if (i != 0 && i != 4 && i != 8 && i != 12 && i != 16 && i != 22)
                {
                    if (i != 0 && dgvPreg3.Rows[i].Cells[2].Value != null)
                    {
                        if (dgvPreg3.Rows[i].Cells[2].Value.ToString() == "Deficiente")
                        {
                            a.U++;
                        }
                        else if (dgvPreg3.Rows[i].Cells[2].Value.ToString() == "Regular")
                        {
                            a.D++;
                        }
                        else if (dgvPreg3.Rows[i].Cells[2].Value.ToString() == "Bueno")
                        {
                            a.T++;
                        }
                        else if (dgvPreg3.Rows[i].Cells[2].Value.ToString() == "Muy Bueno")
                        {
                            a.C++;
                        }
                        else if (dgvPreg3.Rows[i].Cells[2].Value.ToString() == "Excelente")
                        {
                            a.Ci++;
                        }

                        a.totales();

                        lbST13.Text = "" + a.U;
                        lbST23.Text = "" + a.D;
                        lbST33.Text = "" + a.T;
                        lbST43.Text = "" + a.C;
                        lbST53.Text = "" + a.Ci;
                    }
                    else if (i != 0 && i != 4 && i != 8 && i != 12 && i != 16 && i != 22 && dgvPreg3.Rows[i].Cells[2].Value == null)
                    {
                        error = error + "\n" + dgvPreg3.Rows[i].Cells[0].Value;
                    }
                }
            }
            if (error != string.Empty)
            {
                epFail.SetError(this.dgvPreg3, "todavia se debe llenar la pregunta" + error);
            }
            else if (error == string.Empty)
            {
                epFail.Clear();
            }
        }
        private void btnCalcDat4_Click(object sender, EventArgs e)
        {
            error = "";
            lbNC4.Focus();

            a.U = 0;
            a.D = 0;
            a.T = 0;
            a.C = 0;
            a.Ci = 0;


            for (i = 0; i < SDP4.GetLength(0); i++)
            {

                if (i != 0 && i != 4 && i != 8 && i != 12 && i != 16 && i != 22)
                {
                    if (i != 0 && dgvPreg4.Rows[i].Cells[2].Value != null)
                    {
                        if (dgvPreg4.Rows[i].Cells[2].Value.ToString() == "Deficiente")
                        {
                            a.U++;
                        }
                        else if (dgvPreg4.Rows[i].Cells[2].Value.ToString() == "Regular")
                        {
                            a.D++;
                        }
                        else if (dgvPreg4.Rows[i].Cells[2].Value.ToString() == "Bueno")
                        {
                            a.T++;
                        }
                        else if (dgvPreg4.Rows[i].Cells[2].Value.ToString() == "Muy Bueno")
                        {
                            a.C++;
                        }
                        else if (dgvPreg4.Rows[i].Cells[2].Value.ToString() == "Excelente")
                        {
                            a.Ci++;
                        }

                        a.totales();

                        lbST14.Text = "" + a.U;
                        lbST24.Text = "" + a.D;
                        lbST34.Text = "" + a.T;
                        lbST44.Text = "" + a.C;
                        lbST54.Text = "" + a.Ci;
                    }
                    else if (i != 0 && i != 4 && i != 8 && i != 12 && i != 16 && i != 22 && dgvPreg4.Rows[i].Cells[2].Value == null)
                    {
                        error = error + "\n" + dgvPreg4.Rows[i].Cells[0].Value;
                    }
                }
            }
            if (error != string.Empty)
            {
                epFail.SetError(this.dgvPreg4, "todavia se debe llenar la pregunta" + error);
            }
            else if (error == string.Empty)
            {
                epFail.Clear();
            }
        }
        private void btnCalcDat5_Click(object sender, EventArgs e)
        {
            error = "";
            lbNC5.Focus();

            a.U = 0;
            a.D = 0;
            a.T = 0;
            a.C = 0;
            a.Ci = 0;

            for (i = 0; i < SDP5.GetLength(0); i++)
            {

                if (i != 0 && i != 5 && i != 10 && i != 15 && i != 19 && i != 23)
                {
                    if (i != 0 && dgvPreg5.Rows[i].Cells[2].Value != null)
                    {
                        if (dgvPreg5.Rows[i].Cells[2].Value.ToString() == "Deficiente")
                        {
                            a.U++;
                        }
                        else if (dgvPreg5.Rows[i].Cells[2].Value.ToString() == "Regular")
                        {
                            a.D++;
                        }
                        else if (dgvPreg5.Rows[i].Cells[2].Value.ToString() == "Bueno")
                        {
                            a.T++;
                        }
                        else if (dgvPreg5.Rows[i].Cells[2].Value.ToString() == "Muy Bueno")
                        {
                            a.C++;
                        }
                        else if (dgvPreg5.Rows[i].Cells[2].Value.ToString() == "Excelente")
                        {
                            a.Ci++;
                        }

                        a.totales();

                        lbST15.Text = "" + a.U;
                        lbST25.Text = "" + a.D;
                        lbST35.Text = "" + a.T;
                        lbST45.Text = "" + a.C;
                        lbST55.Text = "" + a.Ci;
                    }
                    else if (i != 0 && i != 5 && i != 10 && i != 15 && i != 19 && i != 23 && dgvPreg5.Rows[i].Cells[2].Value == null)
                    {
                        error = error + "\n" + dgvPreg5.Rows[i].Cells[0].Value;
                    }
                }
            }
            if (error != string.Empty)
            {
                epFail.SetError(this.dgvPreg5, "todavia se debe llenar la pregunta" + error);
            }
            else if (error == string.Empty)
            {
                epFail.Clear();
            }
        }
        //dgv's para limpiar errores
        private void dgvPreg1_Click(object sender, EventArgs e)
        {
            epFail.Clear();
        }
        private void dgvPreg2_Click(object sender, EventArgs e)
        {
            epFail.Clear();
        }
        private void dgvPreg3_Click(object sender, EventArgs e)
        {
            epFail.Clear();
        }
        private void dgvPreg4_Click(object sender, EventArgs e)
        {
            epFail.Clear();
        }
        private void dgvPreg5_Click(object sender, EventArgs e)
        {
            epFail.Clear();
        }
        //textbox para nombre
        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {            
            if ((e.KeyChar >= 33 && e.KeyChar <= 64) || (e.KeyChar >= 91 && e.KeyChar <= 96) || (e.KeyChar >= 123 && e.KeyChar <= 126) || (e.KeyChar >= 128 && e.KeyChar <= 255))
            {
                e.Handled = true;
                return;
            }
        }
        private void textBox6_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox6.Text;
            guardado = String.Format(guardado);
            textBox12.Text = guardado;
            textBox18.Text = guardado;
            textBox24.Text = guardado;
            textBox30.Text = guardado;
        }
        private void textBox12_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox12.Text;
            guardado = String.Format(guardado);
            textBox6.Text = guardado;
            textBox18.Text = guardado;
            textBox24.Text = guardado;
            textBox30.Text = guardado;
        }
        private void textBox18_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox18.Text;
            guardado = String.Format(guardado);
            textBox12.Text = guardado;
            textBox6.Text = guardado;
            textBox24.Text = guardado;
            textBox30.Text = guardado;
        }
        private void textBox24_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox24.Text;
            guardado = String.Format(guardado);
            textBox12.Text = guardado;
            textBox18.Text = guardado;
            textBox6.Text = guardado;
            textBox30.Text = guardado;
        }
        private void textBox30_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox30.Text;
            guardado = String.Format(guardado);
            textBox12.Text = guardado;
            textBox18.Text = guardado;
            textBox24.Text = guardado;
            textBox6.Text = guardado;
        }
        //textbox para puesto
        private void textBox5_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox5.Text;
            guardado = String.Format(guardado);
            textBox11.Text = guardado;
            textBox17.Text = guardado;
            textBox23.Text = guardado;
            textBox29.Text = guardado;
        }
        private void textBox11_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox11.Text;
            guardado = String.Format(guardado);
            textBox5.Text = guardado;
            textBox17.Text = guardado;
            textBox23.Text = guardado;
            textBox29.Text = guardado;
        }
        private void textBox17_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox17.Text;
            guardado = String.Format(guardado);
            textBox11.Text = guardado;
            textBox5.Text = guardado;
            textBox23.Text = guardado;
            textBox29.Text = guardado;
        }
        private void textBox23_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox23.Text;
            guardado = String.Format(guardado);
            textBox11.Text = guardado;
            textBox17.Text = guardado;
            textBox5.Text = guardado;
            textBox29.Text = guardado;
        }
        private void textBox29_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox29.Text;
            guardado = String.Format(guardado);
            textBox11.Text = guardado;
            textBox17.Text = guardado;
            textBox23.Text = guardado;
            textBox5.Text = guardado;
        }
        //textbox para unidad 
        private void textBox4_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox4.Text;
            guardado = String.Format(guardado);
            textBox10.Text = guardado;
            textBox16.Text = guardado;
            textBox22.Text = guardado;
            textBox28.Text = guardado;
        }
        private void textBox10_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox10.Text;
            guardado = String.Format(guardado);
            textBox4.Text = guardado;
            textBox16.Text = guardado;
            textBox22.Text = guardado;
            textBox28.Text = guardado;
        }
        private void textBox16_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox16.Text;
            guardado = String.Format(guardado);
            textBox10.Text = guardado;
            textBox4.Text = guardado;
            textBox22.Text = guardado;
            textBox28.Text = guardado;
        }
        private void textBox22_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox22.Text;
            guardado = String.Format(guardado);
            textBox10.Text = guardado;
            textBox16.Text = guardado;
            textBox4.Text = guardado;
            textBox28.Text = guardado;
        }
        private void textBox28_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox28.Text;
            guardado = String.Format(guardado);
            textBox10.Text = guardado;
            textBox16.Text = guardado;
            textBox22.Text = guardado;
            textBox4.Text = guardado;
        }
        //textbox para cod1
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= 32 && e.KeyChar <= 47) || (e.KeyChar >= 58 && e.KeyChar <= 64) || (e.KeyChar >= 91 && e.KeyChar <= 96) || (e.KeyChar >= 123 && e.KeyChar <= 225))
            {
                e.Handled = true;
                return;
            }
        }

        

        private void textBox1_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox1.Text;
            guardado = String.Format(guardado);
            textBox7.Text = guardado;
            textBox13.Text = guardado;
            textBox19.Text = guardado;
            textBox25.Text = guardado;
        }
        private void textBox7_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox7.Text;
            guardado = String.Format(guardado);
            textBox1.Text = guardado;
            textBox13.Text = guardado;
            textBox19.Text = guardado;
            textBox25.Text = guardado;
        }
        private void textBox13_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox13.Text;
            guardado = String.Format(guardado);
            textBox7.Text = guardado;
            textBox1.Text = guardado;
            textBox19.Text = guardado;
            textBox25.Text = guardado;
        }
        private void textBox19_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox19.Text;
            guardado = String.Format(guardado);
            textBox7.Text = guardado;
            textBox13.Text = guardado;
            textBox1.Text = guardado;
            textBox25.Text = guardado;
        }
        private void textBox25_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox25.Text;
            guardado = String.Format(guardado);
            textBox7.Text = guardado;
            textBox13.Text = guardado;
            textBox19.Text = guardado;
            textBox1.Text = guardado;
        }
        //textbox para cod2
        private void textBox2_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox2.Text;
            guardado = String.Format(guardado);
            textBox8.Text = guardado;
            textBox14.Text = guardado;
            textBox20.Text = guardado;
            textBox26.Text = guardado;
        }
        private void textBox8_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox8.Text;
            guardado = String.Format(guardado);
            textBox2.Text = guardado;
            textBox14.Text = guardado;
            textBox20.Text = guardado;
            textBox26.Text = guardado;
        }
        private void textBox14_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox14.Text;
            guardado = String.Format(guardado);
            textBox8.Text = guardado;
            textBox2.Text = guardado;
            textBox20.Text = guardado;
            textBox26.Text = guardado;
        }
        private void textBox20_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox20.Text;
            guardado = String.Format(guardado);
            textBox8.Text = guardado;
            textBox14.Text = guardado;
            textBox2.Text = guardado;
            textBox26.Text = guardado;
        }
        private void textBox26_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox26.Text;
            guardado = String.Format(guardado);
            textBox8.Text = guardado;
            textBox14.Text = guardado;
            textBox20.Text = guardado;
            textBox2.Text = guardado;
        }
        //textbox para cod3
        private void textBox3_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox3.Text;
            guardado = String.Format(guardado);
            textBox9.Text = guardado;
            textBox15.Text = guardado;
            textBox21.Text = guardado;
            textBox27.Text = guardado;
        }
        private void textBox9_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox9.Text;
            guardado = String.Format(guardado);
            textBox3.Text = guardado;
            textBox15.Text = guardado;
            textBox21.Text = guardado;
            textBox27.Text = guardado;
        }
        private void textBox15_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox15.Text;
            guardado = String.Format(guardado);
            textBox9.Text = guardado;
            textBox3.Text = guardado;
            textBox21.Text = guardado;
            textBox27.Text = guardado;
        }
        private void textBox21_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox21.Text;
            guardado = String.Format(guardado);
            textBox9.Text = guardado;
            textBox15.Text = guardado;
            textBox3.Text = guardado;
            textBox27.Text = guardado;
        }
        private void textBox27_KeyUp(object sender, KeyEventArgs e)
        {
            guardado = "";
            guardado = textBox27.Text;
            guardado = String.Format(guardado);
            textBox9.Text = guardado;
            textBox15.Text = guardado;
            textBox21.Text = guardado;
            textBox3.Text = guardado;
        }
    }
}