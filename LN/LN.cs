
//-------------------------------------------------------------------
//====================================================================
// CLASE:  LN
// Lenguaje: C#
// FECHA ULTIMA ACTUALIZACION: 30/Oct/2026 (fechas sin depender de la configuración
//                             regional; VINCheck, NUMERIC, ALF, ALFEng, ALFANUM, FindStr,
//                             DeleteChar y DeleteSTR corregidos)
//====================================================================
using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Net.Mail;
using System.Text.RegularExpressions;   // Using Expresiones Regulares

//=====================================================================
//                        N A M E S P A C E
//=====================================================================
// El programador debe incluir un <using LN.LN;> para 
// acceder a estas funciones y agregar una referencia a 
// la DLL de la clase
//=====================================================================
//-------------------------------------------------------------------
namespace LN
    {
    public class LN
        {
        //-----------------------------------------------------------
        string[] MDA;
        string[] HMS;
        char SPACE = ' ';
        char SLASH = '/';
        string AAAA = "";
        string MM = "";
        string DD = "";
        string FechaNorm = "";
        string HoraNorm = "";
        string FechaSTR = "";
        SmtpClient ServerEMail;
        
        //==============================================================================
        System.Windows.Forms.Timer TimerTemp    = new System.Windows.Forms.Timer();
        Form FormTemp                           = new Form();
        Label LabelTemp                         = new Label();
        System.Drawing.Font FontTemp            = new System.Drawing.Font("Arial",12);
        int TimerCount                          = 0;
        string MensajeTemp                      = "";
        int Segundos                            = 0;
        int CuentaRegresiva                     = 0;
        //------------------------------------------------------------------------------
        public string[] LN_Files;
        public string[] LN_Folders;
        public int LN_NumFiles = 0;
        public int LN_NumFolders = 0;
        //------------------------------------------------------------------------------
        private void TimerTemp_Tick(object sender, EventArgs e)
            {
            if (TimerCount == 0)
                {
                TimerCount++;
                return;
                }
            if (TimerCount == Segundos)
                {
                //----------------------------------------------------------------------
                TimerTemp.Tick -= TimerTemp_Tick; // Desactiva el evento habilitado anteriormente
                TimerTemp.Stop();               // Detiene al timer
                TimerTemp.Dispose();            // Libera los recursos del timer
                TimerTemp.Enabled = false;      // Inhabilita el timer
                //----------------------------------------------------------------------
                FormTemp.Close();
                return;
                }
            TimerCount++;
            CuentaRegresiva--;
            //FormTemp.Text = "Mensaje Temporal: " + CuentaRegresiva;
            return;
            }
        private void SetForm()
            {
            //--------------------------------------------------------------------------
            //System.Drawing.Icon I = new
            //    System.Drawing.Icon(@"D:\____Proyectos VS2015 Community Edition\LN\Annotate_HighPriority.ico");
            // Establece el ancho y el alto del formulario de mensajes
            System.Drawing.Size A = new System.Drawing.Size(700,70);
            //FormTemp.Icon = I;
            /*FormTemp.Text = "Mensaje Temporal: " + CuentaRegresiva; */// Set del texto del mensaje
            FormTemp.Font = FontTemp; // Set del Font a Arial,12
            FormTemp.Height = 60;   // Establece la altura del formulario
            FormTemp.Size = A;    // Setea el ancho y alto del formulario
            // Define la posición del formulario en el centro
            FormTemp.StartPosition = FormStartPosition.CenterScreen;
            // Pone el estado de windows como Normal
            FormTemp.WindowState = FormWindowState.Normal;
            FormTemp.Controls.Add(LabelTemp);   // Adiciona la etiqueta con el mensaje
            //--------------------------------------------------------------------------
            }
        private void SetLabel()
            {
            //--------------------------------------------------------------------------
            System.Drawing.Size A = new System.Drawing.Size(700,70);
            LabelTemp.Size = A;
            LabelTemp.Font = FontTemp;          // Pone el font a Arial,12
            LabelTemp.Text = MensajeTemp;       // Pone el mensaje en el Label
            //--------------------------------------------------------------------------
            }
        private void SetTimer()
            {
            //--------------------------------------------------------------------------
            // *** Setting del Timer ***
            //--------------------------------------------------------------------------
            TimerTemp.Interval = Segundos * 400;    // Establece los segundos a mostrar el mensaje
            TimerTemp.Enabled = true;         // Habilita el Timer
            TimerTemp.Start();                  // Inicia el timer
            TimerTemp.Tick += TimerTemp_Tick;   // Lanza a ejecución el evento Timer
            //--------------------------------------------------------------------------
            }

        //------------------------------------------------------------------------------
        //===========================================================
        // NOMBRE METODO: ExecReport
        //===========================================================
        // PROGRAMADOR:		Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //              object DataSource
        //              string PathReport
        // RETORNO:    NO TIENE
        // TIPO: 	   METODO
        // VERSION:    01
        // FECHA DE INICIO: 	 6/Febrero/2018
        // FECHA DE TERMINACION: 6/Febrero/2018
        //
        // DESCRIPCION:
        // La librería LN posee dos métodos para la emisión de reportes,
        // uno más general y potente (ExecReport), pero algo más complicado, 
        // pues necesita un mayor número de parámetros, y el otro mucho
        // más sencillo y práctico, el cual solamente solo requiere que 
        // se le dé como parámetro, el .RPT generado previamente en forma 
        // autónoma por el Crystal Report en una carpeta cualquiera.
        //
        // Existen dos tipos de reportes en Crystal Report:
        // 1) Los reportes generados totalmente bajo el uso de la
        //    aplicación autónoma del Crystal Report, la cual da
        //    muchas facilidades para una generación rápida de un
        //    reporte a partir de las tablas de una base de datos,
        //    permitiendo agupación de datos (GROUP BY del SQL),
        //    sumarización por columnas incluyendo ruptura por
        //    niveles, por ejemplo, totales por empresas, como nivel
        //    de mayor jerarquía y dentro de este, totalizar por 
        //    departamentos de dicha empresa, y así sucesivamente.
        //    Con la aplicación autónoma del Crystal Report es 
        //    posible generar una gran cantidad de reportes en 
        //    forma rápida al disponer de una interfase muy intuitiva
        //    sin necesidad de utilizar código de programación en C#
        //    para la confección de dichos reportes, por tal razón,
        //    siempre que sea posible le aconsejamos generar el
        //    reporte por esta vía.
        //    Siguiendo la idea anterior podemos crear por ejmplo,
        //    una carpeta que podríamos llamarle D:\REPORTES, donde
        //    se almacenarían todos los reportes .RPT generados en
        //    forma autónoma, y exactamente estos últimos se pueden
        //    mostrar por programa bajo el uso del método ShowRPT 
        //    de nuestra librería LN que lo único que exige a 
        //    diferencia de ExecReport es el path donde se encuentra
        //    el reporte (.RPT)
        // 2) Por otra parte, existen los reportes configurados y 
        //    generados en tiempo objeto, usando código de C#. 
        //    Estos requieren configurarse dinámicamente por programa
        //    y como es lógico, solo pueden ser emitidos por dicho 
        //    programa.
        //
        // EJEMPLO1:
        // Ejemplo de un fragmento de programa de reportes que no se
        // pueden emitir con nuestros dos métodos de LN:
        // . . .
        // // Crear una instancia del objeto de reporte:
        // . . .
        // CrystalReport1 REPORTE = new CrystalReport1();
        //
        // // Establecer el DataSource en cuestión:
        // . . .
        // REPORTE.SetDataSource(DataSetTrabajadores.Tables[0]);
        //
        // Asignar los valores correspondientes a cada parámetro
        // del reporte almacenados en cada variable del programa
        // bajo el uso del método SetParameterValue:
        // . . .
        // REPORTE.SetParameterValue("rpNomTrab",NombreTrabajador);
        // REPORTE.SetParameterValue("rpFechNac",FechaNacimiento);
        // REPORTE.SetParameterValue("rpDir",DireccionTrabajador);
        // REPORTE.SetParameterValue("rpTelf",TelefonoTrabajador);
        // etc.
        // . . .
        // Definir por ejemplo que nuestro reporte se almacenará
        // en un documento Word de Microsoft con el path y el
        // nombre que sigue:
        // PathReport = @"D:\REPORTES\REPORTE.doc";
        // . . .
        // Cuando esté totalmente conformado el reporte pues lo
        // emitimos con el método ExportToDisk del Crystal Report:
        //
        // REPORTE.ExportToDisk(ExportFormatType.WordForWindows,PathReport);
        //
        // Es conveniente aclarar que el primer parámetro del método
        // (DataSource), puede ser del tipo:
        //
        // 1) System.Data.DataTable
        // 2) System.Data.IDataReader
        // 3) IEnumerable DataSource
        // 4) System.Data.DataSet
        //
        // EJEMPLO2:
        // Suponiendo que previamente hemos generado el .RPT con la 
        // aplicación autónoma del Crystal Report en el path que sigue:
        //
        // "D:\REPORTES\ReporteMensual.rpt
        //
        // Pues si queremos que nuestro programa C# muestre el contenido
        // de dicho reporte con el Crystal Viewer, bastaría con poner
        // el siguiente código:
        //
        // LN OBJ = new LN(); // Instanciar un objeto de LN
        // . . .
        // // Mostramos el reporte con el método ShowRPT en
        // // la forma siguiente
        // OBJ.ShowRPT(@"D:\REPORTES\InformeOBAMACARE.rpt");
        //
        // *** ASI DE FÁCIL ***
        //===========================================================
        //private CrystalReportViewer ObjViewer;
        //private Form ObjForm;
        //private ReportDocument ObjReport;
        //---------------------------<Métodos de la Clase>------------------------------
        //private void _ExecReport(object DataSource, string PathReport)
        //    {
        //    //-----------------------------------------------------
        //    // INSTANCIAS NECESARIAS
        //    //-----------------------------------------------------
        //    ObjViewer = new CrystalReportViewer();
        //    ObjForm = new Form();
        //    ObjReport = new ReportDocument();
        //    //-----------------------------------------------------
        //    //     CODIGO EJECUTIVO PARA EMISION DEL REPORTE
        //    //-----------------------------------------------------
        //    ObjForm.Controls.Add(ObjViewer);
        //    ObjViewer.Dock = DockStyle.Fill;
        //    ObjForm.StartPosition = FormStartPosition.CenterScreen;
        //    ObjForm.WindowState = FormWindowState.Maximized;
        //    ObjReport.Load(PathReport);
        //    ObjReport.SetDataSource(DataSource);
        //    ObjViewer.ReportSource = ObjReport;
        //    ObjForm.ShowDialog();
        //    //-----------------------------------------------------
        //    }
        //// Function Overloading para los métodos ExecReport
        //public void ExecReport(System.Data.DataTable DataSource, string PathReport)
        //    {
        //    _ExecReport(DataSource, PathReport);
        //    }
        //public void ExecReport(System.Data.IDataReader DataSource, string PathReport)
        //    {
        //    _ExecReport(DataSource, PathReport);
        //    }
        //public void ExecReport(IEnumerable DataSource, string PathReport)
        //    {
        //    _ExecReport(DataSource, PathReport);
        //    }
        //public void ExecReport(System.Data.DataSet DataSource, string PathReport)
        //    {
        //    _ExecReport(DataSource, PathReport);
        //}
        //============================================================================
        // METODO: ShowReport
        // Fecha de Creación: 26/09/2012
        // Programador: Geovany González y Héctor Fortún
        // Parámetros de Entrada: 
        //              string PathReport,
        // Retorno: No Tiene
        // Descripción: Ya dada anteriormente
        //=====================================================
        //public void ShowRPT(string PathReport)
        //    {
        //    // 1) Instanciación del objeto ObjViewer que llamará al Report Viewer del Crystal Report:
        //    ObjViewer = new CrystalReportViewer();
        //    // 2) Instanciación del objeto ObjForm que mostrará el formulario que contendrá al Report Viewer:    
        //    ObjForm = new Form();
        //    // 3) Instanciación del objeto ObjReport requerido por el Crystal 
        //    ObjReport = new ReportDocument();
        //    // 4) Inserción del objeto ObjViewer dentro del formulario ObjForm
        //    ObjForm.Controls.Add(ObjViewer);
        //    //------------------------<Settings de Propiedades>----------------------------
        //    ObjViewer.Dock = DockStyle.Fill;
        //    ObjForm.StartPosition = FormStartPosition.CenterScreen;
        //    ObjForm.WindowState = FormWindowState.Maximized;
        //    //-----------------------------------------------------------------------------
        //    ObjReport.Load(PathReport);         // Carga el PATH del reporte mediante Load
        //    ObjViewer.ReportSource = ObjReport; // Setea el ReportSource del Viewer
        //    ObjForm.ShowDialog();               // Finalmente muestra el formulario con el reporte
        //    }
        //============================================================================

        public void SendMail(MailMessage MessageParameters)
            {
            ServerEMail.Send(MessageParameters);
            }
        //===========================================================
        // NOMBRE METODO: MailSender	
        //===========================================================
        // PROGRAMADOR:		Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        // string      SMTP,
        // int         PORT,
        // string      CredencialesLogin,
        // string      CredencialesPassword,
        // string      SUBJECT, 
        // string      FROM, 
        // string[]    TO, 
        // string      ATTACH, 
        // string[]    MessageBody)
        // RETORNO:    NO TIENE
        // TIPO: 	   METODO
        // VERSION:    01
        // FECHA DE INICIO: 	 6/Febrero/2018
        // FECHA DE TERMINACION: 6/Febrero/2018
        // DESCRIPCION:
        //===========================================================
        // El método MailSender es el método principal de mensajería
        // de envíos y en el mismo se declaran todos los parámetros
        // necesarios para que el e-mail se transmita correctamente.
        // Los parámetros son los siguientes:
        // string       SMTP:
        //              Aquí se declara la dirección del servidor de
        //              envio (SMTP) que corresponde al cliente de 
        //              mensajería que se desea utilizar. 
        //              Por ejemplo:
        //              "smtp.gmail.com" para correo GMAIL
        //              "smtp.mail.yahoo.com" para correo YAHOO
        //              "smtp.live.com" para correo HOTMAIL
        // int          PORT:
        //              Gmail   ==> Puerto 587 (25 o 465 - pueden presentar problemas)
        //              Yahoo   ==> Puerto 587 (25 o 465 - pueden presentar problemas)
        //              Hotmail ==> Puerto 25 o 465
        //===========================================================
        // NOTA IMPORTANTE PARA CORREO GMAIL (De GOOGLE):
        //
        // El servicio SMTP de Gmail es ofrecido con limitaciones de envíos 
        // para evitar que los  spammers utilicen su servicio para enviar
        // correos basura. 
        //
        // Los límites establecen 100 destinatarios a la vez, y 500 mensajes 
        // por día: si usted supera esta restricción, Google bloqueará su cuenta.
        // Además, su servicio SMTP utiliza IPs compartidas que no garantizan la 
        // entrega, ya que estas también son utilizadas por spammers que dañan 
        // seriamente la reputación de estas direcciones, incrementando las 
        // probabilidades de que su email legítimo sea filtrado por sistemas 
        // antispam por error.
        //
        // Por esta razón, si usted está enviando una campaña de email o necesita 
        // enviar emails ilimitados, la mejor alternativa consiste en optar por un 
        // servidor SMTP profesional: por un precio realmente accesible usted puede 
        // realizar grandes envíos con la certeza de la mayor tasa de envío.
        //===========================================================
        // string      CredencialesLogin,
        // string      CredencialesPassword,
        // string      SUBJECT, 
        // string      FROM, 
        // string[]    TO, 
        // string      ATTACH, 
        // string[]    MessageBody)
        //
        //
        //                               
        //===========================================================
        public void MailSender(string SMTP,
                            int PORT,
                            string CredencialesLogin,
                            string CredencialesPassword,
                            string SUBJECT,
                            string FROM,
                            string[] TO,
                            string ATTACH,
                            string[] MessageBody)
            {
            try
                {
                ServerEMail = new SmtpClient(SMTP, PORT);
                // A continuación se especifica la dirección E-Mail y el Password 
                // correspondiente como parámetros a enviar al server como credenciales
                ServerEMail.Credentials = new
                System.Net.NetworkCredential(CredencialesLogin, CredencialesPassword);
                ServerEMail.EnableSsl = true;
                // Creamos a continuación al objeto OBJMessage que contiene el remitente, los destinatarios, etc

                MailMessage OBJMessage = new MailMessage();

                OBJMessage.Subject = "From HEMA All Insurance";

                int s = 0;
                while (true)
                    {
                    if (TO[s] == null)
                        {
                        break;
                        }
                    OBJMessage.To.Add(TO[s]);
                    s++;
                    }

                // Si se desea que algunos destinatarios reciban copia del mail se debe activar la linea que sigue
                // OBJMessage.CC.Add(new MailAddress(ListaCC));

                OBJMessage.From = new MailAddress(FROM, " <HEMA All Insurances>");

                // A continuación se muestra como atachar un fichero adjunto (Ejemplo)
                if (ATTACH != "")
                    {
                    OBJMessage.Attachments.Add(new Attachment(ATTACH));
                    }

                int k = 0;
                OBJMessage.Body = " ";
                try
                    {
                    while (true)
                        {
                        OBJMessage.Body = OBJMessage.Body + "\n" + MessageBody[k];
                        k++;
                        }
                    }
                catch (Exception E)
                    {
                    if (E.Message != "Índice fuera de los límites de la matriz.")
                        {
                        SendMail(OBJMessage);
                        MessageBox.Show("SEND MAIL - OK", "<* FIN DE ENVIO *>", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        }

                    else
                        {
                        MensERROR("Error Indeterminado en lectura del cuerpo de mensaje - Mensaje No Enviado - Ver a Hector");
                        }
                    return;
                    }

                } // Fin del Try

            catch (Exception E)
                {
                MessageBox.Show(E.ToString());
                }
            } // Fin de MailSender

        //===========================================================
        // NOMBRE METODO: SSNCheck	
        //===========================================================
        // PROGRAMADOR:		Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      string  SSN (Social Security)
        // RETORNO:				bool
        // TIPO: 				METODO
        // VERSION: 			01
        // FECHA DE INICIO: 	 12/Enero/2018
        // FECHA DE TERMINACION: 12/Enero/2018
        // DESCRIPCION:
        //===========================================================
        // Esta función valida la cadena de entrada SSN
        // que representa el Social Security Number
        // de una persona en US.
        // Formato General:
        // 999-99-9999
        //===========================================================

        public bool SSNCheck(string SSN)
            {
            if (Regex.IsMatch(SSN, "^[0-9]{3}-[0-9]{2}-[0-9]{4}$"))
                {
                return true;
                }
            return false;
            } // Fin de SSNCheck(string SSN)

        //===========================================================
        // bool VINCheck(string VIN)
        //===========================================================
        // PARAMETROS DE ENTRADA:	
        // string  VIN (ID del Carro)
        //
        // RETORNO:
        // bool
        // DESCRIPCION:
        // Esta función valida la cadena de entrada VIN
        // que representa el ID de un carro con 17 car.
        //===========================================================

        public bool VINCheck(string VIN)
            {
            // Un VIN tiene siempre 17 caracteres: letras (nunca I, O ni Q)
            // y números. Si no los tiene, no es válido (antes daba error).
            if (VIN == null || !Regex.IsMatch(VIN, "^[A-HJ-NPR-Z0-9]{17}$", RegexOptions.IgnoreCase))
                {
                return false;
                }

            string WMI = "";
            string VDS = "";
            string VIS = "";
            string Car9 = "";

            WMI = VIN.Substring(0, 3);
            VDS = VIN.Substring(3, 6);
            VIS = VIN.Substring(9, 8);
            Car9 = VDS.Substring(5, 1);

            if (Regex.IsMatch(Car9, "[IiOoQqÑñ]"))
                {
                return false; // Es Correcto
                }

            return true;
            } // Fin de VINCheck(string VIN)

        //===========================================================
        // NOMBRE METODO: CheckDateUS_Str
        //===========================================================
        // PROGRAMADOR:				Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //            string Fecha (mm/dd/aaaa)
        // RETORNO:
        // 1 - OK
        // 2 - "dd" no es numérico
        // 3 - "dd" fuera de rango
        // 4 - "mm" no es numérico
        // 5 - "mm" fuera de rango
        // 6 - "aaaa" no es numérico
        // 7 - Primer dígito de¨"aaaa" no es 1 ni 2
        // 8 - Segundo dígito de¨"aaaa" no es 9 ni 0
        // 9 - No tiene slash en posicion 3 o 6
        // 10 - Longitud incorrecta
        // 11- Mes de Febrero con más de 29 días en 
        //     año bisiesto
        // 12- Mes de Febrero con más de 28 días en 
        //     año no bisiesto
        // 13- Los meses de Abril, Junio, Septiembre, 
        //     o Noviembre
        //     no deben tener más de 30 días
        // TIPO: 					int
        // VERSION: 				00
        // FECHA DE INICIO: 		12/Enero/2018
        // FECHA DE TERMINACION: 	12/Enero/2018
        // DESCRIPCION:
        //===========================================================
        // Este método realiza la verificación de una 
        // fecha de formato mm/dd/aaaa
        //===========================================================

        public int CheckDateUS_Str(string Fecha)
            {
            string DD;
            string MM;
            string AAAA;
            int iDD;
            int iMM;
            int iAAAA;
            if (Fecha == null)
                // Longitud de la fecha incorrecta o 
                // formato erroneo
                return 10;
            if (Fecha.Length != 10)
                // Longitud de la fecha incorrecta o 
                // formato erroneo
                return 10;

            MM = Fecha.Substring(0, 2);
            DD = Fecha.Substring(3, 2);
            AAAA = Fecha.Substring(6, 4);

            if (Fecha.Substring(2, 1) != @"/"
                 ||
                 Fecha.Substring(5, 1) != @"/")
                // No tiene slash en posiciones 3 y 6 
                return 9;

            if (!NUM(DD))
                return 2; // "dd" no numerico
            iDD = int.Parse(DD);
            if (iDD < 1
                 ||
                 iDD > 31)
                return 3; // "dd" fuera de rango
            if (!NUM(MM))
                return 4; // "mm" no numerico
            iMM = int.Parse(MM);
            if (iMM < 1
                 ||
                 iMM > 12)
                return 5; // "mm" fuera de rango
            if (!NUM(AAAA))
                return 6; // "aaaa" no numerico
            iAAAA = int.Parse(AAAA);
            if (int.Parse(Fecha.Substring(6, 1)) != 1
                 &&
                 int.Parse(Fecha.Substring(6, 1)) != 2)
                // Primer digito de "aaaa" no es 1 o 2 
                return 7;

            if (int.Parse(Fecha.Substring(7, 1)) != 9
                 &&
                 int.Parse(Fecha.Substring(7, 1)) != 0)
                // Segundo digito de "aaaa" no es 9 o 0 
                return 8;

            if (Bisiesto(AAAA)
                &&
                iMM == 2
                &&
                iDD > 29)
                // Si el año es bisiesto y el mes es
                // Febrero y Tiene más de 29 días, 
                // entonces hay error: 
                // "dd" fuera de rango en año bisiesto
                return 11;

            // Si llega aqui el año no es bisiesto y tiene 
            // solamente 28 días
            if (!Bisiesto(AAAA)
                 &&
                 iMM == 2
                 &&
                 iDD > 28)
                // Si el año no es bisiesto y el
                // mes es Febrero y Tiene más de 28 días,
                // entonces hay error("dd" fuera de rango 
                // en año no bisiesto)
                return 12;

            if ((iMM == 4 ||
                  iMM == 6 ||
                  iMM == 9 ||
                  iMM == 11) &&
                  iDD > 30)
                // Estos meses tienen 
                // un maximo de 30 dias    
                return 13;

            return 1; // No hubo error en Fecha
            } // Fin de CheckDate(string Fecha)  

        //===========================================================
        // NOMBRE METODO: CheckDateESP_Str
        //===========================================================
        // PROGRAMADOR:				Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //            string Fecha (dd/mm/aaaa)
        // RETORNO:
        // 1 - OK
        // 2 - "dd" no es numérico
        // 3 - "dd" fuera de rango
        // 4 - "mm" no es numérico
        // 5 - "mm" fuera de rango
        // 6 - "aaaa" no es numérico
        // 7 - Primer dígito de¨"aaaa" no es 1 ni 2
        // 8 - Segundo dígito de¨"aaaa" no es 9 ni 0
        // 9 - No tiene slash en posicion 3 o 6
        // 10 - Longitud incorrecta
        // 11- Mes de Febrero con más de 29 días en 
        //     año bisiesto
        // 12- Mes de Febrero con más de 28 días en 
        //     año no bisiesto
        // 13- Los meses de Abril, Junio, Septiembre, 
        //     o Noviembre
        //     no deben tener más de 30 días
        // TIPO: 					int
        // VERSION: 				00
        // FECHA DE INICIO: 		05/11/2008
        // FECHA DE TERMINACION: 	05/11/2008
        // DESCRIPCION:
        //===========================================================
        // Este método realiza la verificación de una 
        // fecha de formato dd/mm/aaaa
        //===========================================================

        public int CheckDateESP_Str(string Fecha)
            {
            string DD;
            string MM;
            string AAAA;
            int iDD;
            int iMM;
            int iAAAA;

            if (Fecha.Length != 10)
                // Longitud de la fecha incorrecta o 
                // formato erroneo
                return 10;

            DD = Fecha.Substring(0, 2);
            MM = Fecha.Substring(3, 2);
            AAAA = Fecha.Substring(6, 4);

            if (Fecha.Substring(2, 1) != @"/"
                 ||
                 Fecha.Substring(5, 1) != @"/")
                // No tiene slash en posiciones 3 y 6 
                return 9;

            if (!NUM(DD))
                return 2; // "dd" no numerico
            iDD = int.Parse(DD);
            if (iDD < 1
                 ||
                 iDD > 31)
                return 3; // "dd" fuera de rango
            if (!NUM(MM))
                return 4; // "mm" no numerico
            iMM = int.Parse(MM);
            if (iMM < 1
                 ||
                 iMM > 12)
                return 5; // "mm" fuera de rango
            if (!NUM(AAAA))
                return 6; // "aaaa" no numerico
            iAAAA = int.Parse(AAAA);
            if (int.Parse(Fecha.Substring(6, 1)) != 1
                 &&
                 int.Parse(Fecha.Substring(6, 1)) != 2)
                // Primer digito de "aaaa" no es 1 o 2 
                return 7;

            if (int.Parse(Fecha.Substring(7, 1)) != 9
                 &&
                 int.Parse(Fecha.Substring(7, 1)) != 0)
                // Segundo digito de "aaaa" no es 9 o 0 
                return 8;

            if (Bisiesto(AAAA)
                &&
                iMM == 2
                &&
                iDD > 29)
                // Si el año es bisiesto y el mes es
                // Febrero y Tiene más de 29 días, 
                // entonces hay error: 
                // "dd" fuera de rango en año bisiesto
                return 11;

            // Si llega aqui el año no es bisiesto y tiene 
            // solamente 28 días
            if (!Bisiesto(AAAA)
                 &&
                 iMM == 2
                 &&
                 iDD > 28)
                // Si el año no es bisiesto y el
                // mes es Febrero y Tiene más de 28 días,
                // entonces hay error("dd" fuera de rango 
                // en año no bisiesto)
                return 12;

            if ((iMM == 4 ||
                  iMM == 6 ||
                  iMM == 9 ||
                  iMM == 11) &&
                  iDD > 30)
                // Estos meses tienen 
                // un maximo de 30 dias    
                return 13;

            return 1; // No hubo error en Fecha
            } // Fin de CheckDate(string Fecha)  

        //===========================================================
        // FUNCION NormalizeDateUS_Str
        //===========================================================
        // PROGRAMADOR:		Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      string FechaEntrada
        // RETORNO:				string (CON LA FECHA NORMALIZADA)
        // TIPO: 				METODO
        // VERSION: 			01
        // FECHA DE INICIO: 	 12/Enero/2018
        // FECHA DE TERMINACION: 12/Enero/2018
        // DESCRIPCION:
        //===========================================================
        // Este método normaliza la fecha de entrada que aparece en
        // la variable FechaEntrada en formato US (mm/dd/aaaa) no
        // normalizado a formato US (mm/dd/aaaa) normalizado retornando 
        // a este último. Se considera normalizada la fecha cuando
        // aparece expresado el año con cuatro dígitos, el mes con dos
        // dígitos y el día con dos dígitos.
        // Ejemplo:
        // Entrada: 1/23/1980
        // Salida:  01/23/1980
        //===========================================================
        public string NormalizeDateUS_Str(string FechaEntrada)
            {
            //-------------------------------------------------------
            MDA = FechaEntrada.Split(SPACE);
            FechaEntrada = MDA[0];
            if (FechaEntrada == "")
                {
                return ("0"); // ERROR EN ENTRADA
                }
            if (FechaEntrada.Length > 10 || FechaEntrada.Length < 8)
                {
                return ("0"); // ERROR EN ENTRADA
                }
            MDA = FechaEntrada.Split(SLASH);
            MM = MDA[0];
            DD = MDA[1];
            AAAA = MDA[2];
            //-------------------------------------------------------
            if (!NUM(AAAA) || AAAA.Length != 4)
                {
                return ("0"); // ERROR EN YEAR
                }
            if (AAAA.Length != 4 ||
                 MM.Length < 1 || MM.Length > 2
                 ||
                 DD.Length < 1 || DD.Length > 2)
                {
                return ("0"); // ERROR EN ENTRADA
                }
            //-----------------------------------------------------------
            if (MM.Length == 1)
                {
                MM = "0" + MM;
                }
            if (DD.Length == 1)
                {
                DD = "0" + DD;
                }
            FechaNorm = MM + @"/" + DD + @"/" + AAAA;
            return FechaNorm;
            }

        //===========================================================
        // FUNCION NormalizeDateESP_Str
        //===========================================================
        // PROGRAMADOR:		Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      string FechaEntrada
        // RETORNO:				string
        // TIPO: 				METODO
        // VERSION: 			01
        // FECHA DE INICIO: 	 12/Enero/2018
        // FECHA DE TERMINACION: 12/Enero/2018
        // DESCRIPCION:
        //===========================================================
        // Este método normaliza la fecha de entrada que aparece en
        // la variable FechaEntrada en formato ESP (dd/mm/aaaa) no
        // normalizado a formato ESP (dd/mm/aaaa) normalizado retornando 
        // a este último. Se considera normalizada la fecha cuando
        // aparece expresado el año con cuatro dígitos, el mes con dos
        // dígitos y el día con dos dígitos.
        // Ejemplo:
        // Entrada: 1/1/1980
        // Salida:  01/01/1980
        //===========================================================
        public string NormalizeDateESP_Str(string FechaEntrada)
            {
            //-------------------------------------------------------
            MDA = FechaEntrada.Split(SPACE);
            FechaEntrada = MDA[0];
            if (FechaEntrada == "")
                {
                return ("0"); // ERROR EN ENTRADA
                }
            if (FechaEntrada.Length > 10 || FechaEntrada.Length < 8)
                {
                return ("0"); // ERROR EN ENTRADA
                }
            MDA = FechaEntrada.Split(SLASH);
            DD = MDA[0];
            MM = MDA[1];
            AAAA = MDA[2];
            //------------------------------------------------------
            if (!NUM(AAAA) || AAAA.Length != 4)
                {
                return ("0"); // ERROR EN YEAR
                }
            if (AAAA.Length != 4 ||
                 MM.Length < 1 || MM.Length > 2
                 ||
                 DD.Length < 1 || DD.Length > 2)
                {
                return ("0"); // ERROR EN ENTRADA
                }
            //-----------------------------------------------------------
            if (MM.Length == 1)
                {
                MM = "0" + MM;
                }
            if (DD.Length == 1)
                {
                DD = "0" + DD;
                }
            FechaNorm = DD + @"/" + MM + @"/" + AAAA;
            return FechaNorm;
            }

        //===========================================================
        // FUNCION NormalizeDateISO_Str
        //===========================================================
        // PROGRAMADOR:		Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      string FechaEntrada
        // RETORNO:				string
        // TIPO: 				METODO
        // VERSION: 			01
        // FECHA DE INICIO: 	 12/Enero/2018
        // FECHA DE TERMINACION: 12/Enero/2018
        // DESCRIPCION:
        //===========================================================
        // Este método normaliza la fecha de entrada que aparece en
        // la variable FechaEntrada en formato ISO (aaaa-mm-dd) no
        // normalizado a formato ISO (aaaa-mm-dd) normalizado retornando 
        // a este último. Se considera normalizada la fecha cuando
        // aparece expresado el año con cuatro dígitos, el mes con dos
        // dígitos y el día con dos dígitos.
        // Ejemplo:
        // Entrada: 1980-1-1
        // Salida:  1980-01-01
        //===========================================================
        public string NormalizeDateISO_Str(string FechaEntrada)
            {
            //-------------------------------------------------------
            MDA = FechaEntrada.Split(SPACE);
            FechaEntrada = MDA[0];
            if (FechaEntrada == "")
                {
                return ("0"); // ERROR EN ENTRADA
                }
            if (FechaEntrada.Length > 10 || FechaEntrada.Length < 8)
                {
                return ("0"); // ERROR EN ENTRADA
                }
            MDA = FechaEntrada.Split('-');
            AAAA = MDA[0];
            MM = MDA[1];
            DD = MDA[2];
            //------------------------------------------------------
            if (!NUM(AAAA) || AAAA.Length != 4)
                {
                return ("0"); // ERROR EN YEAR
                }
            if (AAAA.Length != 4 ||
                 MM.Length < 1 || MM.Length > 2
                 ||
                 DD.Length < 1 || DD.Length > 2)
                {
                return ("0"); // ERROR EN ENTRADA
                }
            //-----------------------------------------------------------
            if (MM.Length == 1)
                {
                MM = "0" + MM;
                }
            if (DD.Length == 1)
                {
                DD = "0" + DD;
                }
            //------------------------------------------------------------
            FechaNorm = AAAA.Substring(0, 4) + "-" + MM + "-" + DD;
            return FechaNorm;
            }

        //===========================================================
        // FUNCION ExtractDateUS_Str
        //===========================================================
        // PROGRAMADOR:		Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      DateTime Fecha
        // RETORNO:				string
        // DESCRIPCION:
        //===========================================================
        // Este método extrae la fecha del DateTime con formato US,
        // retornándola normalizada en formato de string.
        //===========================================================
        public string ExtractDateUS_Str(DateTime Fecha)
            {
            //-------------------------------------------------------
            // Formato fijo mm/dd/aaaa, sin depender de la configuración
            // regional de Windows (con Windows en español fallaba).
            //-------------------------------------------------------
            FechaNorm = Fecha.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
            return FechaNorm;
            }
        //===========================================================
        // FUNCION ExtractHour_STR_HM
        //===========================================================
        // PROGRAMADOR:		Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      DateTime Fecha
        // RETORNO:				string
        // DESCRIPCION:
        //===========================================================
        // Este método extrae la fecha del DateTime retornándola
        // normalizada en formato de string.
        //===========================================================
        public string ExtractHour_STR_HM(DateTime Fecha)
            {
            //-------------------------------------------------------
            // Formato fijo hh:mm:ss AM/PM, sin depender de la
            // configuración regional de Windows (con Windows en
            // español fallaba).
            //-------------------------------------------------------
            HoraNorm = Fecha.ToString("hh:mm:ss tt", CultureInfo.InvariantCulture);
            return HoraNorm;
            }
        //===========================================================
        // FUNCION ExtractDateESP_Str
        //===========================================================
        // PROGRAMADOR:		Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      DateTime Fecha
        // RETORNO:				string
        // DESCRIPCION:
        //===========================================================
        // Este método extrae la fecha del DateTime con formato ESP,
        // retornándola normalizada en formato de string.
        //===========================================================
        public string ExtractDateESP_Str(DateTime Fecha)
            {
            //-------------------------------------------------------
            // Formato fijo dd/mm/aaaa, sin depender de la configuración
            // regional de Windows (con Windows en español fallaba y
            // con Windows en inglés devolvía mm/dd/aaaa).
            //-------------------------------------------------------
            FechaNorm = Fecha.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
            return FechaNorm;
            }
        //===========================================================
        // FUNCION NormalizeHour_STR
        //===========================================================
        // PROGRAMADOR:		Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      string FechaEntrada
        // RETORNO:				string (CON LA FECHA NORMALIZADA) o "0"
        //                             si hubo error
        // DESCRIPCION:
        //===========================================================
        // Este método normaliza la fecha de entrada que aparece en
        // la variable FechaEntrada en formato US (mm/dd/aaaa) no
        // normalizado a formato US (mm/dd/aaaa) normalizado retornando 
        // a este último. Se considera normalizada la fecha cuando
        // aparece expresado el año con cuatro dígitos, el mes con dos
        // dígitos y el día con dos dígitos.
        // Ejemplo:
        // Entrada: 1/23/1980
        // Salida:  01/23/1980
        //===========================================================
        public string NormalizeHour_STR(string HoraEntrada)
            {
            string  Hora = "";
            string  Minutos = "";
            string  Segundos = "";
            int     Coincidencias = 0;
            //------------<VALIDACION DE LA ENTRADA>-----------------
            if (HoraEntrada.Length < 5)
                {
                return "0";
                }
            Coincidencias = Regex.Matches(HoraEntrada, @"[:]").Count;
            if (Coincidencias == 0)
                {
                return "0";
                }
            //-------------------------------------------------------
            HMS = HoraEntrada.Split(':');
            Hora = HMS[0];
            Minutos = HMS[1];
            Segundos = HMS[2];
            //-----------------------------------------------------------
            if (Hora.Length == 1)
                {
                Hora = "0" + Hora;
                }
            if (Minutos.Length == 1)
                {
                Minutos = "0" + Minutos;
                }
            if (Segundos.Length == 1)
                {
                Segundos = "0" + Segundos;
                }
            HoraNorm = Hora + ":" + Minutos + ":" + Segundos;
            return HoraNorm;
            }

        //===========================================================
        // FUNCION ConvertDateUStoESP_Str
        //===========================================================
        // PROGRAMADOR:		Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      string FechaEntrada
        // RETORNO:				string
        // TIPO: 				METODO
        // VERSION: 			01
        // FECHA DE INICIO: 	 12/Enero/2018
        // FECHA DE TERMINACION: 12/Enero/2018
        // DESCRIPCION:
        //===========================================================
        // Este método convierte la fecha de entrada que aparece en
        // la variable FechaEntrada en formato US (mm/dd/aaaa) a
        // formato ESP (dd/mm/aaaa) retornando a este último en
        // formato de string.
        // Ejemplo:
        // Entrada: 1/5/1980 (5 de Enero de 1980)
        // Salida:  05/01/1980 
        //===========================================================
        public string ConvertDateUStoESP_Str(string FechaEntrada)
            {
            //-------------------------------------------------------
            MDA = FechaEntrada.Split(SPACE);
            FechaEntrada = MDA[0];
            if (FechaEntrada == "")
                {
                return ("0"); // ERROR EN ENTRADA
                }
            if (FechaEntrada.Length > 10 || FechaEntrada.Length < 8)
                {
                return ("0"); // ERROR EN ENTRADA
                }
            MDA = FechaEntrada.Split(SLASH);
            MM = MDA[0];
            DD = MDA[1];
            AAAA = MDA[2];
            //-----------------------------------------------------------
            if (AAAA.Length != 4 ||
                 MM.Length < 1 || MM.Length > 2
                 ||
                 DD.Length < 1 || DD.Length > 2)
                {
                return ("0"); // ERROR EN ENTRADA
                }
            //-----------------------------------------------------------
            if (MM.Length == 1)
                {
                MM = "0" + MM;
                }
            if (DD.Length == 1)
                {
                DD = "0" + DD;
                }
            FechaNorm = DD + @"/" + MM + @"/" + AAAA.Substring(0, 4);
            return FechaNorm;
            }

        //===========================================================
        // FUNCION ConvertDateESPtoUS_Str
        //===========================================================
        // PROGRAMADOR:		Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      string FechaEntrada
        // RETORNO:				string
        // TIPO: 				METODO
        // VERSION: 			01
        // FECHA DE INICIO: 	 12/Enero/2018
        // FECHA DE TERMINACION: 12/Enero/2018
        // DESCRIPCION:
        //===========================================================
        // Este método convierte la fecha de entrada que aparece en
        // la variable FechaEntrada en formato ESP (dd/mm/aaaa) a
        // formato US (mm/dd/aaaa) retornando a este último en
        // formato de string.
        // Ejemplo:
        // Entrada: 4/5/1980 (04 de Mayo de 1980)
        // Salida:  05/04/1980
        //===========================================================
        public string ConvertDateESPtoUS_Str(string FechaEntrada)
            {
            //-------------------------------------------------------
            MDA = FechaEntrada.Split(SPACE);
            FechaEntrada = MDA[0];
            if (FechaEntrada == "")
                {
                return ("0"); // ERROR EN ENTRADA
                }
            if (FechaEntrada.Length > 10 || FechaEntrada.Length < 8)
                {
                return ("0"); // ERROR EN ENTRADA
                }
            MDA = FechaEntrada.Split(SLASH);
            DD = MDA[0];
            MM = MDA[1];
            AAAA = MDA[2];
            //------------------------------------------------------
            if (MM.Length == 1)
                {
                MM = "0" + MM;
                }
            if (DD.Length == 1)
                {
                DD = "0" + DD;
                }
            FechaNorm = MM + @"/" + DD + @"/" + AAAA.Substring(0, 4);
            return FechaNorm;
            }

        //===========================================================
        // FUNCION ConvertDateISOtoUS_Str
        //===========================================================
        // PROGRAMADOR:		Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      string FechaEntrada
        // RETORNO:				string
        // TIPO: 				METODO
        // VERSION: 			01
        // FECHA DE INICIO: 	 12/Enero/2018
        // FECHA DE TERMINACION: 12/Enero/2018
        // DESCRIPCION:
        //===========================================================
        // Este método convierte la fecha de entrada que aparece en
        // la variable FechaEntrada en formato ISO (aaaa-mm-dd) a
        // formato US (mm/dd/aaaa) retornando a este último en 
        // formato string.
        //===========================================================
        public string ConvertDateISOtoUS_Str(string FechaEntrada)
            {
            if (FechaEntrada == "")
                {
                return (@"01/01/1900");
                }
            MDA = FechaEntrada.Split('-');
            AAAA = MDA[0];
            MM = MDA[1];
            DD = MDA[2];
            if (MM.Length == 1)
                {
                MM = "0" + MM;
                }
            if (DD.Length == 1)
                {
                DD = "0" + DD;
                }
            FechaNorm = MM + @"/" + DD + @"/" + AAAA.Substring(0, 4);
            return FechaNorm;
            }

        //===========================================================
        // FUNCION ConvertDateISOtoESP_Str
        //===========================================================
        // PROGRAMADOR:		Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      string FechaEntrada
        // RETORNO:				string
        // TIPO: 				METODO
        // VERSION: 			01
        // FECHA DE INICIO: 	 12/Enero/2018
        // FECHA DE TERMINACION: 12/Enero/2018
        // DESCRIPCION:
        //===========================================================
        // Este método convierte la fecha de entrada que aparece en
        // la variable FechaEntrada en formato ISO (aaaa-mm-dd) a
        // formato español (dd/mm/aaaa) retornando a este último.
        //===========================================================
        public string ConvertDateISOtoESP_Str(string FechaEntrada)
            {
            //-------------------------------------------------------
            MDA = FechaEntrada.Split(SPACE);
            FechaEntrada = MDA[0];
            if (FechaEntrada == "")
                {
                return ("0"); // ERROR EN ENTRADA
                }
            if (FechaEntrada.Length > 10 || FechaEntrada.Length < 8)
                {
                return ("0"); // ERROR EN ENTRADA
                }
            MDA = FechaEntrada.Split('-');
            AAAA = MDA[0];
            MM = MDA[1];
            DD = MDA[2];
            //------------------------------------------------------
            if (MM.Length == 1)
                {
                MM = "0" + MM;
                }
            if (DD.Length == 1)
                {
                DD = "0" + DD;
                }
            FechaNorm = DD + @"/" + MM + @"/" + AAAA.Substring(0, 4);
            return FechaNorm;
            }
        //===========================================================
        // NOMBRE METODO:  	NUM
        //===========================================================
        // PROGRAMADOR:		Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      string  Numero
        // RETORNO:				bool
        // TIPO: 				METODO
        // VERSION: 			01
        // FECHA DE INICIO: 	 12/Enero/2018
        // FECHA DE TERMINACION: 12/Enero/2018
        // DESCRIPCION:
        //===========================================================
        // Este método chequea si un string es un número entero sin 
        // signo, retornando <true> en caso afirmativo y <false> si 
        // no lo es. 
        //===========================================================
        public bool NUM(string Numero)
            {
            if (Regex.Matches(Numero, "[0-9]").Count == Numero.Length)
                {
                return true;
                }
            return false;
            } // Fin de NUM

        //===========================================================
        // NOMBRE METODO: NUMERIC
        //===========================================================
        // PROGRAMADOR:	Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //             string  Numero
        // RETORNO:
        // true     - Si Numero contiene un número entero o decimal
        //            con signo o sin signo.
        // false    - Si no cumple la condición
        // TIPO: 				METODO
        // VERSION: 			01
        // FECHA DE INICIO: 		12/Enero/2018
        // FECHA DE TERMINACION: 	12/Enero/2018
        // DESCRIPCION:
        //===========================================================
        // Este método chequea si el parámetro de entrada es numérico
        // en cualquiera de sus variantes de representación de enteros
        // o decimales (+11.23, -1.5, 4.2, -.123, +.6932, +321, -4, 
        // 56., -43., 2391, 001, etc.).
        // Retorno:
        // true     - Si se cumple la condición
        // false    - Si no se cumple
        //===========================================================

        public bool NUMERIC(string Numero)
            {
            // Chequeo de los casos con punto al inicio, con o sin signo:
            // Ejemplo: (.15, .972, -.123, +.6932, etc.)
            if (Regex.IsMatch(Numero, @"^[+-]?[.]\d+$"))
                {
                return true;
                }
            // Chequeo de los casos con punto al final sin signo:
            // Ejemplo: (15., 972., etc.)
            if (Regex.IsMatch(Numero, @"^\d+[.]$"))
                {
                return true;
                }
            // Chequeo de los casos con signo y formato {+,-}9(n).9(n):
            // Ejemplo: (+121.151, -45378.97277, etc.)
            if (Regex.IsMatch(Numero, @"^[+-]\d+\.?(\d+)?$"))
                {
                return true;
                }
            // Chequeo de los casos sin signo y formato 9(n).9(n)
            // y también de los enteros sin signo 9(n)
            // Ejemplo: (.15, .972, etc.)
            if (Regex.IsMatch(Numero, @"^\d+\.?(\d+)?$"))
                {
                return true;
                }
            return false;
            } // Fin de NUMERIC 

        //===========================================================
        // NOMBRE METODO:  ZERO
        //===========================================================
        // PROGRAMADOR:	Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      string  Numero
        // RETORNO:
        // true - Si Numero contiene solamente
        //        ceros.
        // false - Caso contrario
        // TIPO: 				METODO
        // VERSION: 			01
        // FECHA DE INICIO: 	 12/Enero/2018
        // FECHA DE TERMINACION: 12/Enero/2018
        // DESCRIPCION:
        //===========================================================
        // Este método chequea si el parámetro de entrada contiene 
        // ceros solamente retornando <true> en caso afirmativo. 
        //===========================================================

        public bool ZERO(string Numero)
            {
            // Chequeo de los casos con formato 0(n)
            if (Regex.IsMatch(Numero, @"^[0]+$"))
                {
                return true;
                }
            return false;
            } // Fin de ZERO 

        //===========================================================
        // NOMBRE METODO:  		ALFANUM
        //===========================================================
        // PROGRAMADOR:	Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //              string  Cadena
        // RETORNO:
        // true - Si la cadena es alfanumérica
        // false - Si no lo es
        // TIPO: 				 bool
        // VERSION: 			 01
        // FECHA DE INICIO: 	 12/Enero/2018
        // FECHA DE TERMINACION: 12/Enero/2018
        // DESCRIPCION:
        //===========================================================
        // Este método chequea si un string es alfanumérico retornando 
        // <true> en caso afirmativo y <false> si no lo es. 
        // Este método considera que una cadena es alfanumérica cuando
        // contiene solo letras y números.  
        //===========================================================
        public bool ALFANUM(string Cadena)
            {
            if (Regex.IsMatch(Cadena, @"^[a-zA-Z0-9]+$"))
                {
                return true;
                }
            return false;
            } // Fin de ALFANUM

        //===========================================================
        // NOMBRE METODO:  	ALF
        //===========================================================
        // PROGRAMADOR:	Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                  string  Cadena
        // RETORNO:	
        // true     - Si la cadena es alfabética
        // false    - Si no lo es 
        // TIPO: 			bool
        // VERSION: 		01
        // FECHA DE INICIO: 	 12/Enero/2018
        // FECHA DE TERMINACION: 12/Enero/2018
        // DESCRIPCION:
        //===========================================================
        // Este método chequea si un string es alfabético bajo el
        // alfabeto español donde existen las letras acentuadas, la ñ,
        // etc., retornando <true> en caso afirmativo y <false> si no lo es.
        // Se considera que una cadena es alfabética en nuestro alfabeto
        // cuando contiene solo letras o aparezcan los caracteres '-',
        // ó apostrofo (') intermedios y que aparezcan una sola vez. 
        // Ejemplo: 
        // Domínguez, O'Donnel Ramírez, Roberto O'Farril Sinforoso
        //===========================================================
        public bool ALF(string Cadena)
            {
            if (Regex.IsMatch(Cadena,
                 @"^[a-zA-ZáéíóúÁÉÍÓÚñÑüÜ\x20]+['-]?[a-zA-ZáéíóúÁÉÍÓÚñÑüÜ\x20]+$"))
                {
                return true;
                }
            return false;
            } // Fin de ALF(string Cadena)

        //===========================================================
        // NOMBRE METODO:  	ALFEng
        //===========================================================
        // PROGRAMADOR:	Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                  string  Cadena
        // RETORNO:	
        // true     - Si la cadena es alfabética
        // false    - Si no lo es 
        // TIPO: 			bool
        // VERSION: 		01
        // FECHA DE INICIO: 	 1/Mayo/2018
        // FECHA DE TERMINACION: 1/Mayo/2018
        // DESCRIPCION:
        //===========================================================
        // Este método chequea si un string es alfabético bajo el
        // alfabeto inglés pero con la característica de que no deben
        // aparecer letras minúsculas, ni acentos ni apóstrofos, ni 
        // guiones tal y como ocurre en la funcion ALF.
        // Ejemplo: 
        // Casos correctos:
        // DOMINGUEZ, ODONNEL, RAMIREZ, ROBERTO, OFARRIL SINFOROSO
        // Casos Incorrectos:
        // Domínguez, ODonnel, Ramirez, Roberto, OFarril Sinforoso
        //===========================================================
        public bool ALFEng(string Cadena)
            {
            if (Regex.IsMatch(Cadena,
                 @"^[a-zA-Z\x20]+['-]?[a-zA-Z\x20]+$"))
                {
                return true;
                }
            return false;
            } // Fin de ALF(string Cadena)
        //===========================================================
        // NOMBRE METODO:  		Bisiesto
        //===========================================================
        // PROGRAMADOR:		Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      string  Cadena
        // RETORNO:				bool
        // TIPO: 				METODO
        // VERSION: 			01
        // FECHA DE INICIO: 	 12/05/2010
        // FECHA DE TERMINACION: 12/05/2010
        // DESCRIPCION:
        //===========================================================
        // Funcion que determina si el año de 
        // entrada es bisiesto, recordar que en 
        // los años bisiestos Febrero tiene 29
        // días y en el resto solo 28
        //===========================================================
        public bool Bisiesto(string AAAAStr)
            {
            int AAAA;

            if (AAAAStr.Length == 0)
                {
                // Retorna false si la longitud 
                // es 0 (Se toma como NO BISIESTO)
                return false;
                }

            AAAA = int.Parse(AAAAStr);

            if (AAAA % 400 == 0)
                {
                // Entonces es bisiesto y 
                // retorno <true>
                return true;
                }

            if (AAAA % 4 == 0 && AAAA % 100 != 0)
                {
                // Entonces es bisiesto y 
                // retorno <true>
                return true;
                }
            // Si se va por aqui no es bisiesto
            return false;
            }

        //===========================================================
        // NOMBRE METODO:  		EdadIntESP
        //===========================================================
        // PROGRAMADOR:		    Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      string FNac
        // RETORNO:
        // -1       ==> Si hubo error de formato en la fecha de entrada
        // edad     ==> Si todo fué OK
        // TIPO: 				METODO
        // VERSION: 			01
        // FECHA DE INICIO: 	 07/01/2019
        // FECHA DE TERMINACION: 07/01/2019
        //===========================================================
        // DESCRIPCION:
        // Este método retorna la edad de una persona hasta la fechacomo 
        // número entero, teniendo como argumento de entrada su Fecha de 
        // Nacimiento en formato de string y regionalización ESN 
        // (español="dd/mm/aaaa"). 
        // Tenga en cuenta que la fecha de entrada tiene que estar 
        // normalizada, es decir, dos dígitos para el día, dos dígitos 
        // para el mes y 4 dígitos para el año.
        //
        // Ejemplo de uso:
        // LN.LN OBJ = new LN.LN(); // Instanciar al objeto OBJ de tipo LN
        //
        // int ALFA = OBJ.EdadStrESP(@"28/06/1948");
        //
        // quedando ALFA = 70 (asumiendo que FNac = "07/01/2019")
        //
        // NOTA IMPORTANTE:
        // Al inicio de la ejecución de este método se asegura que el 
        // formato de entrada sea español de México, el cual coincide 
        // con muchos formatos de fecha del tipo "dd/mm/aaaa" en
        // latinoamerica. Por ejemplo, para Cuba se declara "es-CU",
        // para Argentina "es-AR", etc.
        //
        // Si se desea los valores posibles para el argumento de CultureInfo
        // puede visitar en Internet el sitio:
        // https://msdn.microsoft.com/en-us/library/cc233982.aspx
        // Para Estados Unidos se usa "en-US" donde el formato de fecha
        // es "mm/dd/aaaa".
        //===========================================================
        public
            int EdadIntESP(string FNac)
            {
            double edad;
            string result;
            DateTime FechaDeHoy = DateTime.Now;

            Thread.CurrentThread.CurrentCulture = new CultureInfo("es-MX");
            if (CheckDateESP_Str(FNac) != 1)
                {
                return -1;
                }
            DateTime FechaNacimiento = DateTime.Parse(FNac);
            edad = (FechaDeHoy.
                    Subtract(FechaNacimiento).
                    TotalDays / 365.25);
            result = Trunc(edad, 3, 2);
            return int.Parse(result.Substring(0, 3));
            }
        //===========================================================
        // NOMBRE METODO:  		EdadDecESP
        //===========================================================
        // PROGRAMADOR:		    Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      string FNac
        // RETORNO:
        // -1       ==> Si hubo error de formato en la fecha de entrada
        // edad     ==> Si todo fué OK
        // TIPO: 				METODO
        // VERSION: 			01
        // FECHA DE INICIO: 	 07/01/2019
        // FECHA DE TERMINACION: 07/01/2019
        //===========================================================
        // DESCRIPCION:
        // Este método retorna la edad de una persona hasta la fecha 
        // como número decimal, teniendo como argumento de entrada su 
        // Fecha de Nacimiento en formato de string y regionalización 
        // español = "dd/mm/aaaa". A diferencia del método EdadESP,
        // este da la fecha exacta en forma decimal.
        // Tenga en cuenta que la fecha de entrada tiene que estar 
        // normalizada, es decir, dos dígitos para el día, dos dígitos 
        // para el mes y 4 dígitos para el año.
        //
        // Ejemplo de uso:
        // using LN;
        // ...
        // LN.LN OBJ = new LN.LN(); // Instanciar al objeto OBJ de tipo LN
        //
        // int ALFA = OBJ.EdadStrESP(@"28/06/1948");
        //
        // quedando ALFA = 70 (asumiendo que FNac = "07/01/2019")
        //
        // NOTA IMPORTANTE:
        // Al inicio de la ejecución de este método se asegura que el 
        // formato de entrada sea español de México, el cual coincide 
        // con muchos formatos de fecha del tipo "dd/mm/aaaa" en
        // latinoamerica. Por ejemplo, para Cuba se declara "es-CU",
        // para Argentina "es-AR", etc.
        //
        // Si se desea los valores posibles para el argumento de CultureInfo
        // puede visitar en Internet el sitio:
        // https://msdn.microsoft.com/en-us/library/cc233982.aspx
        //
        // Para Estados Unidos se usa "en-US" donde el formato de fecha
        // es "mm/dd/aaaa".
        //===========================================================
        public
            double EdadDecESP(string FNac)
            {
            double edad;
            DateTime FechaDeHoy = DateTime.Now;

            Thread.CurrentThread.CurrentCulture = new CultureInfo("es-MX");
            if (CheckDateESP_Str(FNac) != 1)
                {
                return -1;
                }
            DateTime FechaNacimiento = DateTime.Parse(FNac);
            edad = (FechaDeHoy.
                    Subtract(FechaNacimiento).
                    TotalDays / 365.25);
            return edad;
            }

        //===========================================================
        // NOMBRE METODO:  		EdadIntUS
        //===========================================================
        // PROGRAMADOR:		    Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      string FNac
        // RETORNO:
        // -1       ==> Si hubo error de formato en la fecha de entrada
        // edad     ==> Si todo fué OK
        // TIPO: 				METODO
        // VERSION: 			01
        // FECHA DE INICIO: 	 07/01/2019
        // FECHA DE TERMINACION: 07/01/2019
        //===========================================================
        // DESCRIPCION:
        // Este método retorna la edad de una persona hasta la fecha como 
        // número entero, teniendo como argumento de entrada su Fecha de 
        // Nacimiento en formato de string y regionalización US 
        // (español="mm/ddd/aaaa"). 
        // Tenga en cuenta que la fecha de entrada tiene que estar 
        // normalizada, es decir, dos dígitos para el día, dos dígitos 
        // para el mes y 4 dígitos para el año.
        //
        // Ejemplo de uso:
        // LN.LN OBJ = new LN.LN(); // Instanciar al objeto OBJ de tipo LN
        //
        // int ALFA = OBJ.EdadStrESP(@"06/28/1948");
        //
        // quedando ALFA = 70 (asumiendo que FNac = "01/07/2019")
        //
        // NOTA IMPORTANTE:
        // Al inicio de la ejecución de este método se asegura que el 
        // formato de entrada sea inglés de US, el cual coincide 
        // con muchos formatos de fecha del tipo "mm/dd/aaaa" en
        // el mundo.
        //
        // Si se desea los valores posibles para el argumento de CultureInfo
        // puede visitar en Internet el sitio:
        // https://msdn.microsoft.com/en-us/library/cc233982.aspx
        //
        //===========================================================
        public
            int EdadIntUS(string FNac)
            {
            double edad;
            string result;
            DateTime FechaDeHoy = DateTime.Now;

            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            if (CheckDateUS_Str(FNac) != 1)
                {
                return -1;
                }
            DateTime FechaNacimiento = DateTime.Parse(FNac);
            edad = (FechaDeHoy.
                    Subtract(FechaNacimiento).
                    TotalDays / 365.25);
            result = Trunc(edad, 3, 2);
            return int.Parse(result.Substring(0, 3));
            }

        //===========================================================
        // NOMBRE METODO:  		EdadDecUS
        //===========================================================
        // PROGRAMADOR:		    Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      string FNac
        // RETORNO:
        // -1       ==> Si hubo error de formato en la fecha de entrada
        // edad     ==> Si todo fué OK
        // TIPO: 				METODO
        // VERSION: 			01
        // FECHA DE INICIO: 	 07/01/2019
        // FECHA DE TERMINACION: 07/01/2019
        //===========================================================
        // DESCRIPCION:
        // Este método retorna la edad de una persona hasta la fecha 
        // como número decimal, teniendo como argumento de entrada su 
        // Fecha de Nacimiento en formato de string y regionalización 
        // US = "mm/dd/aaaa". A diferencia del método EdadUS, este da
        // la fecha exacta en forma decimal.
        // Tenga en cuenta que la fecha de entrada tiene que estar 
        // normalizada, es decir, dos dígitos para el día, dos dígitos 
        // para el mes y 4 dígitos para el año.
        //
        // Ejemplo de uso:
        // using LN;
        // ...
        // LN.LN OBJ = new LN.LN(); // Instanciar al objeto OBJ de tipo LN
        //
        // int ALFA = OBJ.EdadStrESP(@"06/28/1948");
        //
        // quedando ALFA = 70 (asumiendo que FNac = "07/01/2019")
        //
        // NOTA IMPORTANTE:
        // Al inicio de la ejecución de este método se asegura que el 
        // formato de entrada sea inglés de USA, el cual coincide 
        // con muchos formatos de fecha del tipo "mm/dd/aaaa" en
        // el mundo.
        //
        // Si se desea los valores posibles para el argumento de CultureInfo
        // puede visitar en Internet el sitio:
        // https://msdn.microsoft.com/en-us/library/cc233982.aspx
        //
        //===========================================================
        public
            double EdadDecUS(string FNac)
            {
            double edad;
            DateTime FechaDeHoy = DateTime.Now;

            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            if (CheckDateUS_Str(FNac) != 1)
                {
                return -1;
                }
            DateTime FechaNacimiento = DateTime.Parse(FNac);
            edad = (FechaDeHoy.
                    Subtract(FechaNacimiento).
                    TotalDays / 365.25);
            return edad;
            }

        //===========================================================
        // NOMBRE METODO:  		EdadDesdeHastaESP
        //===========================================================
        // PROGRAMADOR:		    Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      string FN, string HT
        // RETORNO:				int
        // TIPO: 				METODO
        // VERSION: 			01
        // FECHA DE INICIO: 	 07/01/2019
        // FECHA DE TERMINACION: 07/01/2019
        //===========================================================
        // DESCRIPCION:
        // Este método Retorna la edad de una persona como entero
        // teniendo como argumentos de entrada su Fecha de Nacimiento 
        // en formato string y la fecha de HASTA en formato string.
        // Es requisito que la fecha (Hasta sea posterior a
        // FechaNacimiento, si esto no se cumple o si existe algún
        // otro error en los parámetros de entrada entonces se retona: -1
        // En este método se asume que las fechas de entrada estan en
        // regionalización español (dd/mm/aaaa)
        //
        // Ejemplo de uso:
        // using LN;
        // ...
        // LN OBJ = new LN();
        // . . .
        // int ALFA = OBJ.EdadDesdeHastaESP(@"28/06/1948", @"17/01/2018");
        // da como resultado:
        // ALFA = 70 
        //===========================================================
        public
            int EdadDesdeHastaESP(string Desde, string Hasta)
            {
            double edad;

            if (CheckDateESP_Str(Desde) != 1 || CheckDateESP_Str(Hasta) != 1)
                {
                return -1;
                }
            DateTime FechaNacimiento = DateTime.Parse(Desde);
            DateTime Hasta1 = DateTime.Parse(Hasta);

            edad = Hasta1.Subtract(FechaNacimiento).TotalDays / 365.25;
            return (int)edad;
            }

        //===========================================================
        // NOMBRE METODO:  		EdadDesdeHastaDecESP
        //===========================================================
        // PROGRAMADOR:		    Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      string FN, string HT
        // RETORNO:				int
        // TIPO: 				METODO
        // VERSION: 			01
        // FECHA DE INICIO: 	 07/01/2019
        // FECHA DE TERMINACION: 07/01/2019
        //===========================================================
        // DESCRIPCION:
        // Este método Retorna la edad exacta de una persona como decimal
        // teniendo como argumentos de entrada su Fecha de Nacimiento 
        // en formato string y la fecha de HASTA en formato string.
        // Es requisito que la fecha (Hasta sea posterior a
        // FechaNacimiento, si esto no se cumple o si existe algún
        // otro error en los parámetros de entrada entonces se retona: -1
        // En este método se asume que las fechas de entrada estan en
        // regionalización español (dd/mm/aaaa)
        //
        // Ejemplo de uso:
        // using LN;
        // ...
        // LN OBJ = new LN();
        // . . .
        // int ALFA = OBJ.EdadDesdeHastaDecimalESP(@"28/06/1948", @"17/01/2018");
        // da como resultado:
        // ALFA = 70.5297...
        //===========================================================
        public
            double EdadDesdeHastaDecESP(string Desde, string Hasta)
            {
            double edad;

            if (CheckDateESP_Str(Desde) != 1 || CheckDateESP_Str(Hasta) != 1)
                {
                return -1;
                }
            DateTime FechaNacimiento = DateTime.Parse(Desde);
            DateTime Hasta1 = DateTime.Parse(Hasta);

            edad = Hasta1.Subtract(FechaNacimiento).TotalDays / 365.25;
            return edad;
            }

        //===========================================================
        // NOMBRE METODO:  		EdadDesdeHastaUS
        //===========================================================
        // PROGRAMADOR:		    Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      string FN, string HT
        // RETORNO:				int
        // TIPO: 				METODO
        // VERSION: 			01
        // FECHA DE INICIO: 	 07/01/2019
        // FECHA DE TERMINACION: 07/01/2019
        //===========================================================
        // DESCRIPCION:
        // Este método Retorna la edad de una persona como entero
        // teniendo como argumentos de entrada su Fecha de Nacimiento 
        // en formato string y la fecha de HASTA en formato string.
        // Es requisito que la fecha (Hasta sea posterior a
        // FechaNacimiento, si esto no se cumple o si existe algún
        // otro error en los parámetros de entrada entonces se retona: -1
        // En este método se asume que las fechas de entrada estan en
        // regionalización US (mm/dd/aaaa)
        //
        // Ejemplo de uso:
        // using LN;
        // ...
        // LN OBJ = new LN();
        // . . .
        // int ALFA = OBJ.EdadDesdeHastaESP(@"06/28/1948", @"01/17/2018");
        // da como resultado:
        // ALFA = 70 
        //===========================================================
        public
            int EdadDesdeHastaUS(string Desde, string Hasta)
            {
            double edad;

            if (CheckDateUS_Str(Desde) != 1 || CheckDateUS_Str(Hasta) != 1)
                {
                return -1;
                }
            DateTime Desde1 = DateTime.Parse(Desde);
            DateTime Hasta1 = DateTime.Parse(Hasta);

            edad = Hasta1.Subtract(Desde1).TotalDays / 365.25;
            return (int)edad;
            }

        //===========================================================
        // NOMBRE METODO:  		EdadDesdeHastaDecUS
        //===========================================================
        // PROGRAMADOR:		    Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      string FN, string HT
        // RETORNO:				int
        // TIPO: 				METODO
        // VERSION: 			01
        // FECHA DE INICIO: 	 07/01/2019
        // FECHA DE TERMINACION: 07/01/2019
        //===========================================================
        // DESCRIPCION:
        // Este método Retorna la edad exacta de una persona como decimal
        // teniendo como argumentos de entrada su Fecha de Nacimiento 
        // en formato string y la fecha de HASTA en formato string.
        // Es requisito que la fecha (Hasta sea posterior a
        // FechaNacimiento, si esto no se cumple o si existe algún
        // otro error en los parámetros de entrada entonces se retona: -1
        // En este método se asume que las fechas de entrada estan en
        // regionalización inglés (mm/dd/aaaa)
        //
        // Ejemplo de uso:
        // using LN;
        // ...
        // LN OBJ = new LN();
        // . . .
        // int ALFA = OBJ.EdadDesdeHastaDecimalUS(@"06/28/1948", @"01/17/2018");
        // da como resultado:
        // ALFA = 70.5297...
        //===========================================================
        public
            double EdadDesdeHastaDecUS(string Desde, string Hasta)
            {
            double edad;

            if (CheckDateESP_Str(Desde) != 1 || CheckDateESP_Str(Hasta) != 1)
                {
                return -1;
                }
            DateTime FechaNacimiento = DateTime.Parse(Desde);
            DateTime Hasta1 = DateTime.Parse(Hasta);

            edad = Hasta1.Subtract(FechaNacimiento).TotalDays / 365.25;
            return edad;
            }

        //--------------------------------------------------
        // METODO: double DateDifD
        //--------------------------------------------------        
        // Método que haya la diferencia en días entre las 
        // fechas DESDE y HASTA (HASTA - DESDE). 
        // Retorno:
        // -1   HASTA no es posterior a DESDE. 
        // -1   Los parametros de entrada no pueden ser nulos. 
        // Un numero decimal mayor o igual que 0 si todo 
        // fue correcto, donde este numero representa la 
        // diferencia en días de HASTA - DESDE. 
        // NOTA:
        // Cuando se quiere introducir una fecha específica
        // para DESDE o HASTA, se puede hacer como mostramos
        // a continuación:
        //
        // DateTime DESDE = new DateTime(2018, 08, 25);
        //
        // Y entonces la fecha quedaría en DESDE bajo formato
        // DateTime de C#.
        //--------------------------------------------------        
        public double DateDifD(DateTime DESDE, DateTime HASTA)
            {
            //----------------------------------------------
            TimeSpan DIF;
            //----------------------------------------------
            double DIAS;
            //-< VALIDACION DE LOS PARAMETROS DE ENTRADA >--
            if (DESDE > HASTA || DESDE == null || HASTA == null)
                return -1; // Hay error
            //---------------------------------------------- 
            // Si llega aqui todo esta correcto
            //----------------------------------------------                              
            DIF = HASTA.Subtract(DESDE);
            DIAS = DIF.TotalDays;
            return DIAS;
            }

        //--------------------------------------------------
        // METODO: double DateDifH
        //--------------------------------------------------        
        // Método que haya la diferencia en horas entre las 
        // fechas DESDE y HASTA (HASTA - DESDE). 
        // Retorno:
        // -1   HASTA no es posterior a DESDE. 
        // -1   Los parametros de entrada no pueden ser nulos. 
        // Un numero decimal mayor o igual que 0 si todo 
        // fue correcto, donde este numero representa la 
        // diferencia en horas de (HASTA - DESDE). 
        // Cuando se quiere introducir una fecha específica
        // para DESDE o HASTA, se puede hacer como mostramos
        // a continuación:
        //
        // DateTime DESDE = new DateTime(2018, 08, 25);
        //
        // Y entonces la fecha quedaría en DESDE bajo formato
        // DateTime de C#.
        //--------------------------------------------------        
        public double DateDifH(DateTime DESDE, DateTime HASTA)
            {
            //----------------------------------------------
            TimeSpan DIF;
            //----------------------------------------------
            double HORAS;
            //-< VALIDACION DE LOS PARAMETROS DE ENTRADA >--
            if (DESDE > HASTA || DESDE == null || HASTA == null)
                return -1; // Hay error
            //---------------------------------------------- 
            // Si llega aqui todo esta correcto
            //----------------------------------------------                              
            DIF = HASTA.Subtract(DESDE);
            HORAS = DIF.TotalHours;
            return HORAS;
            }

        //--------------------------------------------------
        // METODO: double DateDifM
        //--------------------------------------------------        
        // Método que haya la diferencia en meses entre las 
        // fechas DESDE y HASTA (HASTA - DESDE). 
        // Retorno:
        // -1   HASTA no es posterior a DESDE. 
        // -1   Los parametros de entrada no pueden ser nulos. 
        // Un numero decimal mayor o igual que 0 si todo 
        // fue correcto, donde este numero representa la 
        // diferencia en meses de HASTA - DESDE. 
        //
        // NOTA IMPORTANTE:
        // Es conveniente aclarar que dado que la cantidad de
        // días de un mes es variable e incluso puede variar
        // en Febrero en los años bisiestos, pues el cálculo
        // exacto se convierte en un verdadero algoritmo, por
        // tal razón en este método se asume que cada mes tiene
        // 30 días, lo cual es frecuentemente utilizado en
        // muchos proyectos a fin de simplificar la solución.
        // Por esta razón si Ud. desea una mayor precisión
        // tendrá que hacerlo tratando el número de días de 
        // cada mes en particular a fín de obtener el número
        // exacto de meses en la diferencia de fechas.
        //
        // SUGERENCIA:
        // Cuando se quiere introducir una fecha específica
        // para DESDE o HASTA, se puede hacer como mostramos
        // a continuación:
        //
        // DateTime DESDE = new DateTime(2018, 08, 25);
        //
        // Y entonces la fecha quedaría en DESDE bajo formato
        // DateTime de C#.
        //--------------------------------------------------        
        public double DateDifM(DateTime DESDE, DateTime HASTA)
            {
            //----------------------------------------------
            TimeSpan DIF;
            //----------------------------------------------
            double DIAS;
            //-< VALIDACION DE LOS PARAMETROS DE ENTRADA >--
            if (DESDE > HASTA || DESDE == null || HASTA == null)
                return -1; // Hay error
            //---------------------------------------------- 
            // Si llega aqui todo esta correcto
            //----------------------------------------------                              
            DIF = HASTA.Subtract(DESDE);
            DIAS = DIF.TotalDays;

            return (DIAS / 30.0);
            }

        //==========================================================
        // NormalizeDate
        //==========================================================
        // Descripción:
        // Esta función realiza la normalización de las cantidades 
        // dadas en los primeros tres parámetros de modo tal que si
        // la cantidad de meses de un año es mayor que 12 o si la 
        // cantidad de días es mayor que 30 se realiza el ajuste
        // necesario quedando el resultado en las variables de
        // salida ANorm, MNorm, y DNorm. La transmisión a estos
        // últimos se hace por referencia y de esta forma se puede
        // retornar más de un valor.
        //
        // Ejemplo:
        // using LN;
        // ...
        // LN.LN OBJ = new LN.LN();
        // int AA = 0;
        // int MM = 135;
        // int DD = 105; 
        // bool Result;
        // ...
        // Result = OBJ.NormalizeDatesESP(0,135,105,AA,BB,DD)
        //                                
        // Retorno:
        // true  - Si todo fué OK
        //         En este caso quedaría:
        //         AA = 
        //         BB =
        //         CC =
        // false - Si hubo error
        //========================================================
        public bool NormalizeDate(int A,
                                  int M,
                                  int D,
                                  ref int ANorm,
                                  ref int MNorm,
                                  ref int DNorm)
            {
            //--------------------------------------------------
            int IntAnos;
            int CocienteAnos;
            int IntMeses;
            int CocienteMeses;
            int IntDias;
            //--------------------------------------------------
            double COCIENTE;
            //--------------------------------------------------
            if (D <= 0)
                {
                return false;
                }
            //--------------------------------------------------
            IntAnos = A;
            IntMeses = M;
            IntDias = D;
            //---------------------<DIAS>-----------------------
            // Se comienza dividiendo la cantidad total de dias
            // de entrada por 30 y esto nos dara la cantidad de
            // meses contenidos en ese total de dias, y despues
            // se toma el resto de dicha division que ya contiene
            // la cantidad real de dias despues de normalizado
            //--------------------------------------------------
            COCIENTE = (D / 30);
            // El cociente son meses
            CocienteMeses = (int)COCIENTE;
            DNorm = (D % 30); // El resto son dias
            //---------------------<MESES>----------------------
            // Antes de proceder a la normalizacion de los meses
            // sumamos el cociente de la division anterior a
            // IntMeses.
            //--------------------------------------------------
            M = M + CocienteMeses;
            COCIENTE = (M / 12);
            CocienteAnos = (int)COCIENTE; // El cociente son años
            MNorm = (M % 12); // El resto son meses
            //---------------------<AÑOS>----------------------- 
            ANorm = A + CocienteAnos;
            //--------------------------------------------------
            return true;
            }

        //===========================================================
        // NOMBRE METODO:  		Porciento
        //===========================================================
        // PROGRAMADOR:		    Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      double percent
        //                      double num
        // RETORNO:				double
        // TIPO: 				METODO
        // VERSION: 			01
        // FECHA DE INICIO: 	 23/01/2018
        // FECHA DE TERMINACION: 23/01/2018
        //===========================================================
        // DESCRIPCION:
        // Este método Retorna el porciento P del numero N.
        // Por ejemplo, si queremos saber el 5% de 20 invocaríamos a
        // la función de la forma siguiente: 
        // double result = OBJ.Porciento(5,20)
        // Y result = 1
        //===========================================================
        public double Porciento(double Percent, double Number)
            {
            return (Number * Percent / 100);
            }
        //===========================================================
        // NOMBRE METODO:  		MSDOSWait
        //===========================================================
        // PROGRAMADOR:		    Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      string comando
        // RETORNO:				Código de terminación del proceso
        //                      invocado (ExitCode)
        // TIPO: 				METODO
        // VERSION: 			01
        // FECHA DE INICIO: 	 18/05/2009
        // FECHA DE TERMINACION: 18/05/2009
        //===========================================================
        // DESCRIPCION:
        // Este método ejecuta un comando MS-DOS en una ventana de 
        // consola quedando en espera de la ejecución del comando
        // para continuar sin cerrar la ventana de consola  en espera
        // de que termine la ejecución del comando invocado.
        //===========================================================
        public int MSDOSWait(string comando)
            {
            Process P = new Process();
            P.StartInfo.FileName = "CMD.EXE";
            P.StartInfo.Arguments = "/K " + comando;
            P.Start();
            P.WaitForExit();
            return P.ExitCode;
            }
        //===========================================================
        // NOMBRE METODO:  		MSDOSNoWait
        //===========================================================
        // PROGRAMADOR:		    Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      string comando
        // RETORNO:				Código de terminación del proceso
        //                      invocado (ExitCode)
        // TIPO: 				METODO
        // VERSION: 			01
        // FECHA DE INICIO: 	 18/05/2009
        // FECHA DE TERMINACION: 18/05/2009
        //===========================================================
        // DESCRIPCION:
        // Este método ejecuta un comando MS-DOS en una ventana de 
        // consola pero no queda en espera como el caso anterior.
        // Por esta razón la tarea se ejecuta asincrónicamente respecto
        // al programa llamador.
        //===========================================================
        public int MSDOSNoWait(string comando)
            {
            Process P = new Process();
            P.StartInfo.FileName = "CMD.EXE";
            P.StartInfo.Arguments = "/C " +
                                    comando;
            P.Start();
            return P.ExitCode;
            }
        //===========================================================
        // NOMBRE METODO:  		PAUSE()
        //===========================================================
        // PROGRAMADOR:		    Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      string comando
        // RETORNO:				NO TIENE
        // TIPO: 				METODO
        // VERSION: 			01
        // FECHA DE INICIO: 	 18/05/2009
        // FECHA DE TERMINACION: 18/05/2009
        //===========================================================
        // DESCRIPCION:
        // Provoca una pausa en la consola de comandos, (PAUSE) del
        // MS-DOS emitiendo el siguiente mensaje:
        // *** PRESIONE CUALQUIER TECLA PARA CONTINUAR ***
        //===========================================================
        public
            void PAUSE()
            {
            Console.
            WriteLine(
                  "*** PRESIONE CUALQUIER TECLA PARA CONTINUAR ***"
                     );
            Console.
            ReadLine();
            }
        //=======================================================
        // NOMBRE METODO:  		ExecProcess
        //=======================================================
        // PROGRAMADOR:			Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      string comando
        // RETORNO:				Código de terminación.
        // TIPO: 				METODO
        // VERSION: 			01
        // FECHA DE INICIO: 		03/07/2009
        // FECHA DE TERMINACION: 	03/07/2009
        // DESCRIPCION:
        //=======================================================
        // Descripción:
        // Este método permite invocar cualquier fichero de cualquier
        // extensión reconocida por Windows, por tanto, si invocamos
        // a un fichero .docx, pues el Windows abre al mismo como
        // fichero Word de Office, así como cualquier otra extensión
        // reconocida por el sistema.
        // Con este es también [posible llamar a cualquier tipo de 
        // ejecutable (.exe, .bat, .msi, msc, etc.)
        // A continuación  mostramos tres ejemplos de llamadas.
        // Ejemplo1: 
        // Abrir al documento "ReporteMensual.docx" que se encuentra 
        // dentro de la carpeta CARPET de D:\
        //
        // OBJLN.ExecProcess(@"D:\CARPET\ReporteMensual.docx");
        //
        // Ejemplo2: 
        // (Ejecución de un .bat que se encuentra en la raiz de D:\
        //
        // OBJLN.ExecProcess(@"D:\ABC.bat");
        //
        // Ejemplo3: 
        // (Ejecución de un .exe que aparece  dentro  de una carpeta 
        // cuyo nombre contiene espacios, en este caso el PATH debe
        // estar entre comillas y siguiendo  la  sintáxis de C# por
        // cada comilla que  se  desee  insertar  en  el  string de 
        // especificación hay que poner dos, por esta razón aparecen
        // tres comillas al inicio y al final. Si Ud. no hace esto
        // el ejecutable referenciado no ejecutará correctamente.
        //
        // OBJLN.ExecProcess(@"""C:\Renamer 2.20\ReNamer.exe"""); 
        //=======================================================
        public int ExecProcess(string comando)
            {
            Process P = new Process();
            P.StartInfo.FileName = comando;
            P.Start();
            P.WaitForExit();
            return P.ExitCode;
            }
        //=======================================================
        // NOMBRE METODO:  		ExecProcessNoWait
        //=======================================================
        // PROGRAMADOR:			Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      string comando
        // RETORNO:				El código de retorno 
        //                      del proceso lanzado
        // TIPO: 				METODO
        // VERSION: 			01
        // FECHA DE INICIO: 		03/07/2009
        // FECHA DE TERMINACION: 	03/07/2009
        // DESCRIPCION:
        //=======================================================
        // Descripción:
        // Este método es similar al ExecProcess pero 
        // no espera a la terminación del proceso 
        // lanzado y continua la ejecución de la 
        // siguiente instrucción.
        //=======================================================
        public int ExecProcessNoWait(string comando)
            {
            Process P = new Process();
            P.StartInfo.FileName = comando;
            P.Start();
            return P.ExitCode;
            }

        //==============================================================
        // Funcion Trunc(double Origen,
        //               int Enteros,
        //               int Decimales);
        //==============================================================
        // Descripción:
        // Esta función toma el contenido double de la variable Origen
        // y lo repersenta con la cantidad de cifras de la variable Enteros
        // para la parte entera y la cantidad de cifras de Decimales para
        // la parte decimal.
        // Retorno:
        // Un string que contiene el valor double de origen y lo 
        // representa en el número de dígitos que contiene Enteros
        // y el número de dígitos que contiene Decimales.
        // La parte entera se rellena con ceros a la izquierda si es
        // necesario.
        //=============================================================
        // Ejemplo:
        // Si Origen = 12.33444
        //    Enteros = 5
        //    Decimales = 2
        // Entonces retorna un string = "00012.33"
        //=============================================================
        public string Trunc(double Origen,
                          int Enteros,
                          int Decimales)
            {
            string Mascara;
            string Destino = "";

            Mascara = FillString("0", Enteros) +
                      "." +
                      FillString("0", Decimales + 1);
            Destino = Origen.ToString(Mascara).
                      Substring(0, Mascara.Length - 1);
            return Destino;
            }

        //==============================================================
        // Funcion BetweenInc(int N, int a, int b);
        //==============================================================
        // Descripción:
        // Método que determina si N se encuentra en el intervalo [a,b]
        // incluyendo los extremos
        // Retorno:
        // true     - Si se encuentra en el intervalo
        // false    - Si no se encuentra en el intervalo 
        //=============================================================
        public bool BetweenInc(int N, int a, int b)
            {
            if (N >= a && N <= b)
                {
                return true;
                }
            return false;
            }
        //==============================================================
        // Funcion BetweenIncDec(double N, double a, double b); (Sobrecargada)
        //==============================================================
        // Descripción:
        // Método que determina si N se encuentra en el intervalo [a,b]
        // incluyendo los extremos
        // Retorno:
        // true     - Si se encuentra en el intervalo
        // false    - Si no se encuentra en el intervalo 
        //=============================================================
        public bool BetweenIncDec(double N, double a, double b)
            {
            if (N >= a && N <= b)
                {
                return true;
                }
            return false;
            }
        //==============================================================
        // Funcion BetweenExc(int N, int a, int b);
        //==============================================================
        // Descripción:
        // Método que determina si N se encuentra en el intervalo (a,b)
        // excluyendo los extremos
        // Retorno:
        // true     - Si se encuentra en el intervalo
        // false    - Si no se encuentra en el intervalo 
        //=============================================================
        public bool BetweenExc(int N, int a, int b)
            {
            if (N > a && N < b)
                {
                return true;
                }
            return false;
            }
        //==============================================================
        // Funcion BetweenExc(double N, double a, double b); (Sobrecargada)
        //==============================================================
        // Descripción:
        // Método que determina si N se encuentra en el intervalo (a,b)
        // excluyendo los extremos
        // Retorno:
        // true     - Si se encuentra en el intervalo
        // false    - Si no se encuentra en el intervalo 
        //=============================================================
        public bool BetweenExcDec(double N, double a, double b)
            {
            if (N > a && N < b)
                {
                return true;
                }
            return false;
            }
        //=======================================================
        // NOMBRE METODO:  		MensTemp
        //=======================================================
        // PROGRAMADOR:	Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //   string M
        //   int    Seg
        // RETORNO:	NO TIENE
        // FECHA DE TERMINACION: 	02/12/2019
        //=======================================================
        // DESCRIPCION:
        //=======================================================
        // El método MensTemp muestra un mensaje temporal de modo
        // tal que al cumplirse el tiempo transcurrido, dicho 
        // mensaje desaparece y continua la ejecución del programa.
        // El parámetro de entrada M de tipo string, contiene el
        // mensaje a mostrar, mientras que el segundo parámetro
        // Seg especifíca el número de segundos que permanecerá
        // el mensaje en pantalla.
        //=======================================================
        public void MensTemp(string M, int Seg)
            {
            //--------------------<Inicialización de Variables>-------------------------
            TimerCount = 0;
            MensajeTemp = M;
            Segundos = Seg;
            CuentaRegresiva = Seg;
            //--------------------------------------------------------------------------
            if (M.Length > 60)
                {
                MensERROR("Longitud del Mensaje en Función MsgTemp No Debe Ser Mayor a 60");
                return;
                }
            //---------------------------------------------------------------------------
            SetLabel();
            SetForm();
            SetTimer();
            //---------------------------------------------------------------------------
            FormTemp.ShowDialog(); // Muestra el formulario del mensaje
            }
        //=====================================================
        // Nombre de la Función: MensINFO(string Mensaje)
        // Fecha de Creación: 30/04/2012
        // Programador: Héctor Fortún
        // Parámetros de Entrada: Mensaje
        // Retorno: No Tiene
        // Descripción:
        // Emite un mensaje informativo con el texto que aparece
        // en el parámetro de entrada
        //=====================================================
        public void MensINFO(string Mensaje)
            {
            MessageBox.Show(Mensaje,
                            "<INFORMATIVO>",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
            }
        //=====================================================
        // Nombre de la Función: Menserror(string Mensaje)
        // Fecha de Creación: 30/04/2012
        // Programador: Héctor Fortún
        // Parámetros de Entrada: Mensaje
        // Retorno: No Tiene
        // Descripción:
        // Emite un mensaje de error con el texto que aparece
        // en el parámetro de entrada
        //=====================================================
        public void MensERROR(string Mensaje)
            {
            MessageBox.Show(Mensaje,
                            "<ERROR>",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
            }
        //=====================================================
        // Nombre de la Función: MensConfirmacion(string Mensaje)
        // Fecha de Creación: 30/04/2012
        // Programador: Héctor Fortún
        // Parámetros de Entrada: Mensaje
        // Retorno: No Tiene
        // Descripción:
        // Emite un mensaje de petición de confirmación cuyo texto
        // se encuentra en el parámetro de entrada
        //=====================================================
        public bool MensConfirm(string Mensaje)
            {
            DialogResult result;
            result = MessageBox.Show(Mensaje,
                                     "<Confirmación>",
                                     MessageBoxButtons.YesNo,
                                     MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
                {
                return true;
                }
            else
                {
                return false;
                }
            }
        //=====================================================
        // METODO: Upper
        //=====================================================
        // Fecha de Creación: 09/01/2019
        // Programador: Héctor Fortún
        // Parámetros de Entrada: Cadena
        // Retorno: No Tiene
        //=====================================================
        // DESCRIPCION:
        //=====================================================
        // Convierte a mayúsculas la cadena del parámetro de
        // entrada. 
        // La transmisión de parámetros es por referencia.
        // Ejemplo:
        // LN OBJ = new LN();
        // string CAD = "abc";
        // ...
        // OBJ.Upper(ref CAD); 
        // Después de la ejecución del método Upper quedaría:
        // CAD = "ABC"
        //=====================================================
        public void Upper(ref string Cadena)
            {
            Cadena = Cadena.ToUpper();
            }
        //=====================================================
        // METODO: Lower
        //=====================================================
        // Fecha de Creación: 09/01/2019
        // Programador: Héctor Fortún
        // Parámetros de Entrada: Cadena
        // Retorno: No Tiene
        //=====================================================
        // DESCRIPCION:
        //=====================================================
        // Convierte a minúsculas la cadena del parámetro de
        // entrada. 
        // La transmisión de parámetros es por referencia.
        // Ejemplo:
        // LN OBJ = new LN();
        // string CAD = "ABC";
        // ...
        // OBJ.Lower(ref CAD); 
        // Después de la ejecución del método Lower quedaría:
        // CAD = "abc"
        //=====================================================
        public void Lower(ref string Cadena)
            {
            Cadena = Cadena.ToLower();
            }
        //===========================================================
        // NOMBRE METODO: FindStr	
        //===========================================================
        // PROGRAMADOR:		Lic. Hector Fortun
        // PARAMETROS DE ENTRADA:	
        //                      string Cadena
        //                      string Subcadena
        // RETORNO:				int
        // TIPO: 				METODO
        // VERSION: 			01
        // FECHA DE INICIO: 	 15/Febrero/2018
        // FECHA DE TERMINACION: 15/Febrero/2018
        // DESCRIPCION:
        //===========================================================
        // Esta función analiza si Cadena contiene al menos una 
        // Subcadena retornando el número de apariciones de esta 
        // última dentro de Cadena.
        //===========================================================

        public int FindStr(string Cadena, string Subcadena)
            {
            if (string.IsNullOrEmpty(Subcadena))
                return 0;
            // Regex.Escape: la Subcadena se busca tal cual (".", "(" o "+"
            // no tienen significado especial)
            return Regex.Matches(Cadena, Regex.Escape(Subcadena)).Count;
            }

        //===========================================================
        // public string DeleteChar(string Cadena, string CHAR)
        //===========================================================
        // PARAMETROS DE ENTRADA:	
        // string Cadena
        // string CHAR
        // RETORNO:
        // string
        // DESCRIPCION:
        //===========================================================
        // Esta función borra cualquier aparición de los caracteres 
        // especificados en CHAR, dentro de Cadena, eliminándolos y
        // retornando a la cadena resultante del proceso.
        // Ejemplo:
        // Si CAD = "(786)234-9856" y CAR = "()-"
        // entonces la función retorna:
        // 7862349856
        // Nota Importante:
        // Los caracteres en CAR pueden estar en cualquier orden y el
        // resultado sería el mismo, es decir, si CAR = ")-(", pues 
        // tendríamos el mismo resultado.
        //===========================================================
        public string DeleteChar(string Cadena, string CHAR)
            {
            // Se quita cada carácter que aparezca en CHAR, en cualquier orden
            string CadenaReemplazada = string.Concat(Cadena.Where(C => !CHAR.Contains(C)));
            return CadenaReemplazada;
            }

        //===========================================================
        // public string DeleteSTR(string Cadena, string SubCadena)
        //===========================================================
        // PARAMETROS DE ENTRADA:	
        // string Cadena
        // string Subcadena
        // RETORNO:
        // string
        // DESCRIPCION:
        //===========================================================
        // Esta función borra cualquier aparición de SubCadena, 
        // dentro de Cadena, eliminándola y retornando a la cadena 
        // resultante del proceso.
        // Ejemplo:
        // Si CAD = "123abc456abc789" y CAR = "abc"
        // entonces la función retorna:
        // "123456789"
        // A diferencia de la función DeleteChar, aquí si importa el
        // orden de aparición de los caracteres de SubCadena, pues la
        // búsqueda no es de los caracteres "a", "b" o "c" por separado
        // sino de "abc" con su orden de aparición.
        //===========================================================

        public string DeleteSTR(string Cadena, string SubCadena)
            {
            if (string.IsNullOrEmpty(SubCadena))
                return Cadena;
            // Se busca SubCadena tal cual (".", "(" o "+" no tienen significado especial)
            string CadenaReemplazada = Cadena.Replace(SubCadena, "");
            return CadenaReemplazada;
            }

        //==============================================================
        // Funcion FillString(string Cadena, string Car, int NVeces);
        //==============================================================
        // Descripción:
        // Deposita en <Cadena> al caracter <Car> N veces. La cadena es
        // borrada previamente a la propagación.
        // Retorno: La Cadena conteniendo a Car N veces.
        //=============================================================
        public string FillString(string Car, int NVeces)
            {
            string CADENA = "";
            for (int i = 0; i < NVeces; i++)
                {
                CADENA = CADENA + Car;
                }
            return CADENA;
            }

        //=====================================================
        // METODO: SplitString
        //=====================================================
        // Fecha de Creación: 10/08/2012
        // Programador: Héctor Fortún
        // Parámetros de Entrada: 
        //              ref string[] Resultado, 
        //              ref delim    Cadena, 
        //              string[]     delimitador
        // Retorno: No Tiene
        //=====================================================
        // DESCRIPCION
        //=====================================================
        // Extrae las subcadenas separadas por el delimitador 
        // especificado en el tercer parámetro del string que
        // se encuentra contenido en Cadena, dejándo las subcadenas
        // en el arreglo de tipo string Resultado, donde la
        // transmisión es por referencia como se observa
        //=====================================================
        // Ejemplos de Uso:
        // Caso1:
        // En la secuencia que aparece a continuación, se toma como único
        // delimitador la coma.
        //
        // string[] RESULT = {""}; // <= La parte derecha inicializa el arreglo
        //                         //    RESULT, y esto es obligatorio antes de
        //                         //    usar la función SplitString
        // string[] comas = new string[] {","}; // Define el separador como coma
        // string source = "UNO,DOS,TRES,CUATRO";
        // OBJ.SplitString(ref RESULT, source, comas);
        //
        // En este caso quedaría en el array RESULT lo siguiente:
        // RESULT[0] = "UNO";
        // RESULT[1} = "DOS";
        // RESULT[2} = "TRES";
        // RESULT[3} = "CUATRO";
        //
        // Caso 2:
        // En la secuencia que aparece a continuación, se toman varios
        // delimitadores, la coma, un punto y un punto y coma.
        // string[] delim = new string[] {",",".",";"};
        // string source = "UNO,DOS.TRES;CUATRO";
        // OBJ.SplitString(ref RESULT, source, delim);
        // En este caso quedaría en el array RESULT lo siguiente:
        // RESULT[0} = "UNO";
        // RESULT[1} = "DOS";
        // RESULT[2} = "TRES";
        // RESULT[3} = "CUATRO";
        //=====================================================
        public void SplitString(ref string[] Resultado,
                                string Cadena,
                                string[] delimitador)
            {
            Resultado = Cadena.Split(delimitador, StringSplitOptions.None);
            }

        //=============================================================================================
        //                                 Metodo FileDiskOpenCreate()
        //=============================================================================================
        // FORMATO:
        // public StreamWriter FileDiskOpenCreate(string Path)

        // DESCRIPCION:
        // Este método recibe como entrada una cadena de caracteres que describe el directorio donde
        // vamos a crear el fichero descrito en Path. 
        // Si dicho directorio contiene sub-directorios, estos no se obtienen por este método ya que
        // para este fin se dispone del método:
        // FoldersDiskOpenCreate()
        // que se describe más adelante.

        // PARAMETROS DE ENTRADA:
        // string Path (contiene el full path del fichero a crear).
        // Nota: Recomendamos que use @"path_del_fichero" a fin de usar los backslash (\) normales,
        // sin necesidad de ponerlos duplicados en cada aparición.
        // Ejemplo: 
        // . . .
        // StreamWriter FileOut;
        // . . .
        // FileOut = FileDiskOpenCreate(@"D:\Empleados\Departamento1\Dep1.txt");

        // RETORNO:
        // 1) Retorna null si se presenta algún error en la apertura del fichero.
        // 2) Retorna un elemento del tipo StreamWriter que debe ser referenciado por el método
        //    FileDiskWrite
        //    que es el encargado de escribir un artículo en el fichero seleccionado y que se describe    
        //    más adelante.
        //=============================================================================================
        public StreamWriter FileDiskOpenCreate(string Path)
            {
            StreamWriter OutFile;
            try
                {
                Stream OutFolder = new FileStream(Path, FileMode.Create, FileAccess.Write);
                OutFile = new StreamWriter(OutFolder);
                }
            catch
                {
                return null; // Error en Creación del Fichero
                }
            return OutFile; // Creación del Fichero OK
            }

        //=============================================================================================
        //                                  Metodo FileDiskOpenRead
        //=============================================================================================
        // FORMATO:
        // public StreamReader FileDiskOpenRead(string Path)

        // DESCRIPCION:
        // Este método recibe como entrada una cadena de caracteres que describe el directorio de donde
        // vamos a leer el fichero descrito en Path. 

        // PARAMETROS DE ENTRADA:
        // string Path (contiene el full path del fichero desde donde se va a leer).
        // Nota:
        // Recomendamos que use @"path_del_fichero" a fin de usar los backslash (\) normales, sin
        // necesidad de ponerlos duplicados en cada aparición.

        // Ejemplo: 
        // . . .
        // StreamReader FileIn;
        // . . .
        // FileIn = FileDiskOpenRead(@"D:\Empleados\Departamento1\Dep1.txt");

        // RETORNO:
        // 1) Retorna null si se presenta algún error en la apertura del fichero.
        // 2) Retorna un elemento del tipo StreamReader que debe ser referenciado por el método
        //     FileDiskRead
        //     que es el encargado de leer un artículo en el fichero seleccionado y que se describe    
        //     más adelante.
        //=============================================================================================
        public StreamReader FileDiskOpenRead(string Path)
            {
            StreamReader InFile;
            try
                {
                Stream InFolder = new FileStream(Path, FileMode.Open, FileAccess.Read);
                InFile = new StreamReader(InFolder);
                }
            catch
                {
                return null; // Error en Creación del Fichero
                }
            return InFile; // Creación del Fichero OK
            }

        //=============================================================================================
        //                                Metodo FileDiskOpenAppend()
        //=============================================================================================
        // FORMATO:
        // public StreamWriter FileDiskOpenAppend(string Path)

        // DESCRIPCION:
        // Este método recibe como entrada una cadena de caracteres que describe el directorio que
        // contiene el fichero donde vamos a adicionar los nuevos artículos mediante el método 
        // FileDiskWrite que se describe más adelante.
        // Si dicho directorio contiene sub-directorios, estos no se obtienen por este método ya que
        // para este fin se dispone del método:
        //
        // FoldersDiskOpenCreate()
        // que se describe más adelante.

        // PARAMETROS DE ENTRADA:
        // string Path (contiene el full path del fichero al que se van a adicionar nuevos artículos).
        // Nota:
        // Recomendamos que use @"path_del_fichero" a fin de usar los backslash (\) normales,
        // sin necesidad de ponerlos duplicados en cada aparición.
        // Ejemplo: 
        // . . .
        // StreamWriter FileOut;
        // . . .
        // FileOut = FileDiskOpenCreate(@"D:\Empleados\Departamento1\Dep1.txt");

        // RETORNO:
        // 1) Retorna null si se presenta algún error en la apertura del fichero.
        // 2) Retorna un elemento del tipo StreamWriter que debe ser referenciado por el método
        //     FileDiskWrite
        //     que es el encargado de escribir/adicionar un nuevo artículo en el fichero seleccionado y 
        //     que se describe más adelante.
        //=============================================================================================
        public StreamWriter FileDiskOpenAppend(string Path)
            {
            StreamWriter OutFile;
            try
                {
                Stream OutFolder = new FileStream(Path, FileMode.Append, FileAccess.Write);
                OutFile = new StreamWriter(OutFolder);
                }
            catch
                {
                return null; // Error en Creación del Fichero
                }
            return OutFile; // Creación del Fichero OK
            }

        //=============================================================================================
        //                                  Metodo FileDiskWrite
        //=============================================================================================
        // FORMATO:
        // public bool FileDiskWrite(StreamWriter OutFile, string Line)

        // DESCRIPCION:
        // Este método realiza la escritura del artículo contenido en Line en dependencia de como fué
        // abierto dicho fichero (Create o Append).

        // PARAMETROS DE ENTRADA:
        // 1) StreamWriter OutFile: 
        // Aquí en OutFile se consigna la variable del tipo de datos StreamWriter que se obtuvo en el
        // retorno de alguno de los métodos FileDiskOpenCreate o FileDiskOpenAppend.
        // 2) string Line:
        // La variable Line contiene el artículo en formato de string que se desea escribir en el
        // fichero abierto.
        //
        // Ejemplo: 
        // . . .
        // string Artículo = "";
        // StreamWriter FileOut;
        // . . .
        // FileOut = FileDiskOpenCreate(@"D:\Empleados\Departamento1\Dep1.txt");
        // if	(FileOut == null)
        //      {
        //      *** Secuencia de tratamiento de error
        //      }
        // . . .
        // Artículo = "Expediente=015,Nombre=Leslie Sullivan";
        // FileDiskWrite(FileOut, Artículo);  // Escribe el artículo anterior en el fichero Dep1.txt
        // . . .
        // . . .
        // RETORNO:
        // 1) Retorna false si se presenta algún error en la escritura del artículo.
        // 2) Retorna true si todo fué OK.
        //=============================================================================================
        public bool FileDiskWrite(StreamWriter OutFile, string Line)
            {
            bool retorno = true;
            try
                {
                OutFile.WriteLine(Line);
                }
            catch
                {
                retorno = false; // Error en Escritura de Artículo
                }
            return retorno; // Escritura de Artículo OK
            }

        //=============================================================================================
        //                                  Metodo FileDiskRead()
        //=============================================================================================
        // FORMATO:
        // public string FileDiskRead(StreamReader InFile)
        // DESCRIPCION:
        // Este método realiza la lectura del artículo en curso retornándolo en formato string.
        // Este método esta concebido para ficheros .TXT (no aplica para ficheros binarios).
        //
        // PARAMETROS DE ENTRADA:
        // 1) StreamReader InFile: 
        // Aquí en InFile se consigna la variable del tipo de datos StreamReader que se obtuvo en el 
        // retorno del método FileDiskOpenRead.
        // Ejemplo: 
        // . . .
        // string Artículo = "";
        // StreamReader FileIn;
        // . . .
        // FileIn = FileDiskOpenRead(@"D:\Empleados\Departamento1\Dep1.txt");
        // if	(FileIn == null)
        //      {
        //      *** Secuencia de tratamiento de error
        //      }
        // . . .
        // Recibe el artículo en curso dentro del fichero Dep1.txt depositándolo en Artículo 
        // Artículo = FileDiskRead(FileIn);  
        // . . .
        // . . .
        // RETORNO:
        // 1) Retorna null si se presenta algún error en la escritura del artículo.
        // 2) Retorna el artículo en curso si todo fué OK.
        //=============================================================================================
        public string FileDiskRead(StreamReader InFile)
            {
            string Line = "";

            try
                {
                Line = InFile.ReadLine();
                }
            catch
                {
                return null; // Error en Escritura de Artículo
                }
            return Line; // Escritura de Artículo OK
            }

        //=============================================================================================
        //                                Metodo FileDiskCloseWriter()
        //=============================================================================================
        // FORMATO:
        // public bool FileDiskCloseWriter(StreamWriter OutFile)

        // DESCRIPCION:
        // Este método realiza el cierre del fichero definido en FileDiskOpenCreate o
        // FileDiskOpenAppend
        // PARAMETROS DE ENTRADA:
        // 1) StreamWriter OutFile: 
        // Aquí en OutFile se consigna la variable del tipo de datos StreamWriter que se obtuvo en el
        // retorno de alguno de los métodos FileDiskOpenCreate o FileDiskOpenAppend.
        // Ejemplo: 
        // . . .
        // string Artículo = "";
        // StreamWriter FileOut;
        // . . .
        // FileOut = FileDiskOpenCreate(@"D:\Empleados\Departamento1\Dep1.txt");
        // if	(FileOut == null)
        // {
        // *** Secuencia de tratamiento de error
        // }
        // . . .
        // Artículo = "Expediente=015,Nombre=Leslie Sullivan";
        // FileDiskWrite(FileOut, Artículo);  // Escribe el artículo anterior en el fichero Dep1.txt
        // . . .
        // . . .
        // FileDiskCloseWriter(FileOut);
        // . . .

        // RETORNO:
        // 1) Retorna false si se presenta algún error en la escritura del artículo.

        // 2) Retorna true si todo fué OK.
        //=============================================================================================
        public bool FileDiskCloseWriter(StreamWriter OutFile)
            {
            bool retorno = true;
            try
                {
                OutFile.Close();
                }
            catch
                {
                retorno = false; // Error en Close del Fichero de Salida
                }
            return retorno; // Close OK
            }

        //=============================================================================================
        //                               Metodo FileDiskCloseReader()
        //=============================================================================================
        //FORMATO:
        //public bool FileDiskCloseReader(StreamReader IntFile)

        //DESCRIPCION:
        //Este método realiza el cierre del fichero definido en FileDiskOpenRead.

        //PARAMETROS DE ENTRADA:
        //1) StreamReader InFile: 
        //Aquí en InFile se consigna la variable del tipo de datos StreamReader que se obtuvo en el 
        //retorno del método FileDiskOpenRead.
        //Ejemplo: 
        //. . .
        //string Artículo = "";
        //StreamReader FileIn;
        //. . .
        //FileIn = FileDiskOpenReader(@"D:\Empleados\Departamento1\Dep1.txt");
        //if	(FileIn == null)
        //      {
        //      *** Secuencia de tratamiento de error
        //      }
        //. . .
        // // Obtiene el artículo en curso del fichero Dep1.txt y lo deposita en Artículo
        //Artículo = FileDiskRead(FileIn); 
        //. . .
        //. . .
        //FileDiskCloseReader(FileOut);
        //. . .
        //RETORNO:
        //1) Retorna false si se presenta algún error en la escritura del artículo.
        //2) Retorna true si todo fué OK.
        //=============================================================================================
        public bool FileDiskCloseReader(StreamReader IntFile)
            {
            bool retorno = true;
            try
                {
                IntFile.Close();
                }
            catch
                {
                retorno = false; // Error en Close del Fichero de Salida
                }
            return retorno; // Close OK
            }

        //======================================================================================
        //                              GetDirectoryFilesDialog()
        //======================================================================================
        //             Seleccionar el directorio a listar con diálogo de entrada
        //======================================================================================
        // FORMATO:
        // public string[] GetDirectoryFilesDialog(ref int NumeroDeFicheros)

        // DESCRIPCION:
        // Este método retorna la lista de ficheros seleccionados dentro del diálogo en un arreglo de
        //         tipo string[] dejando también el número de los mismos en la variable NumeroDeFicheros
        // (transmisión por referencia por lo que ref es necesario).

        // PARAMETROS DE ENTRADA:
        // ref int NumeroDeFicheros
        // Variable transmitida por referencia que obtiene el número de ficheros contenidos dentro del directorio seleccionado en el diálogo. Si el directorio seleccionado contiene subcarpetas, estas no se cuentan por este método pues para este fin se tiene a los métodos:
        // GetDirectoryFoldersDialog y GetDirectoryFoldersPath.
        // Ejemplo: 
        // . . .
        // int NumeroDeFicheros = 0;
        // . . .
        // string ListaDeFicheros[] = GetDirectoryFilesDialog(ref NumeroDeFicheros); 
        // . . . 
        // *Nota:
        // En este ejemplo la lista de ficheros obtenida se copia en el arreglo ListaDeFicheros[] definido por el programador, quedando la cantidad de los mismos en la variable entera del segundo parámetro NumeroDeFicheros.
        // RETORNO:
        // 1) Retorna null si se presenta algún error en la apertura del fichero.
        // 2) Retorna  la lista de ficheros que pertenecen al directorio seleccionado  en 
        // ListaDeFicheros[] 
        // y en el segundo parámetro (NumeroDeFicheros) la cantidad de ficheros.
        //======================================================================================
        public string[] GetDirectoryFilesDialog(ref int NumeroDeFicheros)
            {
            LN_Files = null;
            LN_NumFiles = 0;
            //------------------<DECLARACION DE VARIABLES>----------------------
            string Retorno = "";
            FolderBrowserDialog Folder1 = new FolderBrowserDialog();
            string FolderName1 = "";
            //------------------------------------------------------------------
            Retorno = Folder1.ShowDialog().ToString(); // Muestro Dialogo de Seleccion
            if (Retorno != "OK")
                {
                return null;
                }
            try
                {
                FolderName1 = Folder1.SelectedPath;
                DirectoryInfo D = new DirectoryInfo(FolderName1);

                // Obtengo la lista de Ficheros y Subdirectorios del Directorio Seleccionado
                LN_Files = Directory.GetFiles(FolderName1);
                LN_NumFiles = D.GetFiles().Length;

                // Preparación de la función SplitString
                string[] Backslash = new string[] { @"\" };
                string[] RESULT = {""};
                string[] RESULTDir = {""};

                for (int i = 0; i < LN_NumFiles; i++)
                    {
                    //--------------------< Llenar los arrays de Files>---------------------
                    SplitString(ref RESULT, LN_Files[i], Backslash);
                    LN_Files[i] = RESULT[RESULT.Length - 1];
                    }
                // Retorna la lista de  contenido en el directorio seleccionado y el
                // Numero de Ficheros del Directorio
                NumeroDeFicheros = LN_NumFiles;
                return LN_Files;
                }
            catch
                {
                return null;
                }
            }

        //===================================================================================
        //                             GetDirectoryFilesPath()
        //     Seleccionar el directorio a listar con Path como parámetro de entrada
        //===================================================================================
        // FORMATO:
        // public string[] GetDirectoryFilesPath(string Path, ref int NumeroDeFicheros)
        // DESCRIPCION:
        // Este método retorna la lista de ficheros contenidos en el directorio descrito en Path 
        // en un arreglo de tipo string[] y el número de los mismos en la variable del segundo 
        // parámetro NumeroDeFicheros (transmisión por referencia por lo que ref es necesario).
        // PARAMETROS DE ENTRADA:
        // string Path
        // Es un string que define el Path del directorio del cual queremos obtener su lista de ficheros
        // ref int NumeroDeFicheros
        // Variable transmitida por referencia que obtiene el número de ficheros del directorio escogido.
        // Ejemplo: 
        // . . .
        // int NumeroDeFicheros = 0;
        // . . .
        // string ListaDeFicheros[] = GetDirectoryFilesPath(@"D:\Empleados\Dep\Dep1.txt",
        //         ref NumeroDeFicheros); 
        // . . . 
        // *Nota:
        // En este ejemplo la lista de ficheros obtenida se copia en el arreglo ListaDeFicheros[] definido por el programador, quedando la cantidad de los mismos en la variable entera del segundo parámetro NumeroDeFicheros.
        // RETORNO:
        // 1) Retorna null si se presenta algún error en la apertura del fichero.
        // 2) Retorna  la lista de ficheros que pertenecen al directorio seleccionado y en el segundo parámetro (NumeroDeFicheros) la cantidad de ficheros de dicha lista.
        //===================================================================================
        public string[] GetDirectoryFilesPath(string Path, ref int NumeroDeFicheros)
            {
            LN_Files = null;
            LN_NumFiles = 0;
            //------------------<DECLARACION DE VARIABLES>----------------------
            string FolderName1 = "";
            //------------------------------------------------------------------
            if (Path == "" || Path == null)
                {
                return null;
                }
            try
                {
                FolderName1 = Path;
                DirectoryInfo D = new DirectoryInfo(FolderName1);

                // Obtengo la lista de Ficheros y Subdirectorios del Directorio Seleccionado
                LN_Files = Directory.GetFiles(FolderName1);
                LN_NumFiles = D.GetFiles().Length;

                // Preparación de la función SplitString
                string[] Backslash = new string[] { @"\" };
                string[] RESULT = {""};
                string[] RESULTDir = {""};

                for (int i = 0; i < LN_NumFiles; i++)
                    {
                    //--------------------< Llenar los arrays de Files>---------------------
                    SplitString(ref RESULT, LN_Files[i], Backslash);
                    LN_Files[i] = RESULT[RESULT.Length - 1];
                    }
                // Retorna la lista de  contenido en el directorio seleccionado y el
                // Numero de Ficheros del Directorio
                NumeroDeFicheros = LN_NumFiles;
                return LN_Files;
                }
            catch
                {
                return null;
                }
            }

        //=========================================================================================
        // GetDirectoryFoldersDialog() - Obtener subdirectorios dentro del directorio seleccionado
        //=========================================================================================
        //FORMATO:

        //public string[] GetDirectoryFoldersDialog(ref int NumeroDeSubDirectorios)

        //DESCRIPCION:
        //Este método obtiene la lista de subfolders contenidos dentro del directorio escogido
        //por el usuario, retornando el resultado en un arreglo de tipo string[] y depositándo
        //también el número de subfolders en la variable del parámetro NumeroDeSubDirectorios 
        //(transmisión por referencia por lo que ref es necesario).

        //PARAMETROS DE ENTRADA:
        //ref int NumeroDeSubDirectorios
        //Aqui queda después de la ejecución del método el número de subdirectorios contenidos
        //en la carpeta seleccionada.
        //Ejemplo: 
        //. . .
        //int NumeroDeFolders = 0;
        //string ListaSubFolders[] = null;
        //. . .
        //ListaSubFolders[] = GetDirectoryFoldersDialog(ref NumeroDeFolders); 
        //. . . 
        //*Nota:
        // En este ejemplo la lista de subfolders obtenida se copia en el arreglo
        // ListaSubFolders[] definido por el programador, quedando la cantidad de
        // los mismos en la variable entera del parámetro NumeroDeFolders transmitida
        // por referencia.

        //RETORNO:
        //1) Retorna null si se presenta algún error en la apertura del fichero.
        //2) Retorna  la lista de folders que pertenecen al directorio seleccionado,
        //   y en el parámetro NumeroDeFolders la cantidad de los mismos.
        //====================================================================================
        public string[] GetDirectoryFoldersDialog(ref int NumeroDeSubDirectorios)
            {
            //------------------<DECLARACION DE VARIABLES>----------------------
            string Retorno = "";
            FolderBrowserDialog Folder1 = new FolderBrowserDialog();
            string FolderName1 = "";
            //------------------------------------------------------------------
            Retorno = Folder1.ShowDialog().ToString(); // Muestro Dialogo de Seleccion
            if (Retorno != "OK")
                {
                return null;
                }
            try
                {
                FolderName1 = Folder1.SelectedPath;
                DirectoryInfo D = new DirectoryInfo(FolderName1);
                DirectoryInfo D2 = new DirectoryInfo(FolderName1);

                // Obtengo la lista de Ficheros y Subdirectorios del Directorio Seleccionado
                LN_Folders = Directory.GetDirectories(FolderName1);
                LN_NumFolders = D2.GetDirectories().Length;

                // Preparación de la función SplitString
                string[] Backslash = new string[] { @"\" };
                string[] RESULT = {""};
                string[] RESULTDir = {""};

                if (LN_NumFolders != 0) // Es que hay subcarpetas
                    {
                    for (int i = 0; i < LN_NumFolders; i++)
                        {
                        SplitString(ref RESULTDir, LN_Folders[i], Backslash);
                        LN_Folders[i] = RESULTDir[RESULTDir.Length - 1];
                        }
                    }
                // Retorna la lista de  contenido en el directorio seleccionado y el
                // Numero de Subdirectorios del Directorio
                NumeroDeSubDirectorios = LN_NumFolders;
                return LN_Folders;
                }
            catch
                {
                return null;
                }
            }

        //===========================================================================================
        // string[] GetDirectoryFoldersPath() - Obtener subdirectorios dentro del directorio
        //                                      seleccionado con Path
        //===========================================================================================
        //FORMATO:
        //public string[] GetDirectoryFoldersPath(string Path, ref int NumeroDeSubDirectorios)

        //DESCRIPCION:
        //Este método obtiene la lista de subfolders contenidos dentro del directorio escogido por
        // el usuario en Path, retornando el resultado en un arreglo de tipo string[] y depositándo
        // también el número de subfolders de dicha lista en la variable del segundo parámetro

        //        ref int NumeroDeSubDirectorios

        //(transmisión por referencia por lo que ref es necesario).

        //PARAMETROS DE ENTRADA:
        //string Path
        //Describe el Path del directorio que se desea listar.
        //ref int NumeroDeSubDirectorios
        //Aquí queda el número de subdirectorios del directorio seleccionado en Path

        //Ejemplo: 
        //. . .
        //int NumeroDeFolders = 0;
        //string ListaSubFolders[] = null;
        //. . .
        //ListaSubFolders[] = GetDirectoryFoldersPath(@"D:\Empleados\Dep\Dep1.txt",
        //                                            ref NumeroDeFolders); 
        //. . . 
        //*Nota:
        //En este ejemplo la lista de subfolders obtenida se copia en el arreglo ListaSubFolders[]
        //definido por el programador, quedando la cantidad de los mismos en la variable entera del 
        //segundo parámetro NumeroDeFolders.

        //RETORNO:
        //1) Retorna null si se presenta algún error en la apertura del fichero.
        //2) Retorna  la lista de folders que pertenecen al directorio seleccionado y en el segundo
        //   parámetro la cantidad de los mismos.
        //===========================================================================================
        public string[] GetDirectoryFoldersPath(string Path, ref int NumeroDeSubDirectorios)
            {
            //------------------<DECLARACION DE VARIABLES>----------------------
            string FolderName1 = "";
            //------------------------------------------------------------------
            if (Path == "" || Path == null)
                {
                return null;
                }
            try
                {
                FolderName1 = Path;
                DirectoryInfo D = new DirectoryInfo(FolderName1);
                DirectoryInfo D2 = new DirectoryInfo(FolderName1);

                // Obtengo la lista de Ficheros y Subdirectorios del Directorio Seleccionado
                LN_Folders = Directory.GetDirectories(FolderName1);
                LN_NumFolders = D2.GetDirectories().Length;

                // Preparación de la función SplitString
                string[] Backslash = new string[] { @"\" };
                string[] RESULT = {""};
                string[] RESULTDir = {""};

                if (LN_NumFolders != 0) // Es que hay subcarpetas
                    {
                    for (int i = 0; i < LN_NumFolders; i++)
                        {
                        SplitString(ref RESULTDir, LN_Folders[i], Backslash);
                        LN_Folders[i] = RESULTDir[RESULTDir.Length - 1];
                        }
                    }
                // Retorna la lista de  contenido en el directorio seleccionado y el
                // Numero de Subdirectorios
                NumeroDeSubDirectorios = LN_NumFolders;
                return LN_Folders;
                }
            catch
                {
                return null;
                }
            }

        //=============================================================================================
        //                                  Metodo CreateFileParameter
        //=============================================================================================
        // FORMATO:
        // public bool CreateFileParameter(string PathFile, string Parameter)
        // DESCRIPCION:
        // Este método escribe la línea de parámetros contenida en Parameter dentro del fichero
        //         descrito en PathFile que será creado.
        // PARAMETROS DE ENTRADA:
        // 1) string PathFile
        //     Contiene el Path del fichero de parámetros que será creado.
        // 2) string Parameter
        //     Contiene la línea de parámetros que será escrita en el fichero dado.
        // Ejemplo: 
        // . . .
        // CreateFileParameter(@"C:\Param.txt", "ID=1020, Name=Rose, Age=25");
        // . . . 
        // RETORNO:
        // 1) Retorna false si se presenta algún error en la apertura del fichero.
        // 2) Retorna true si todo fué OK.
        //=============================================================================================
        public bool CreateFileParameter(string PathFile, string Parameter)
            {
            //------------------<DECLARACION DE VARIABLES>----------------------
            StreamWriter ParameterFile;
            //------------------------------------------------------------------
            if (PathFile == "" || PathFile == null)
                {
                return false;
                }
            try
                {
                ParameterFile = FileDiskOpenCreate(PathFile);
                FileDiskWrite(ParameterFile, Parameter);
                FileDiskCloseWriter(ParameterFile);
                }
            catch
                {
                return false; // Error en Creación del Fichero
                }
            return true; // Creación del Fichero OK
            }

        //=============================================================================================
        //                               Metodo CreateFileParameterArray
        //=============================================================================================
        //FORMATO:

        //public bool CreateFileParameterArray(string PathFile, string[] Parameter)

        //DESCRIPCION:
        //Este método escribe las líneas de parámetros contenidas en el array Parameter 
        //del segundo operando en el fichero descrito en PathFile que será creado.

        //PARAMETROS DE ENTRADA:

        //1) string PathFile
        //    Describe el Path del fichero de parámetros que será creado.

        //2) string[] Parameter
        //    Array que contiene las líneas de parámetros que serán escritas en el fichero
        //        dado.

        //Ejemplo: 
        //. . .
        //string[] Lineas = {"Parametro1","Parametro2","Parametro3"};
        //CreateFileParameterArray(@"D:\Param.txt", Lineas);
        //. . . 

        //En este ejemplo se crea el fichero C:\Param.txt y se escriben los artículos:
        //Parametro1
        //Parametro2
        //Parametro3

        //RETORNO:

        //1) Retorna false si se presenta algún error en la apertura del fichero.

        //2) Retorna true si todo fué OK.
        //Nota:
        //En Windows 10 el programador tiene que tener habilitado los derechos de escritura 
        //en la raiz del disco C:\ si desea crear su fichero parámetro en este lugar, pues 
        //usualmente el C:\ es el disco sistema y hay un mayor nivel de protección por parte 
        //del sistema. Por esta razón aconsejamos que utilice cualquier otro disco para este 
        //fin.
        //=============================================================================================

        public bool CreateFileParameterArray(string PathFile, string[] Parameter)
            {
            //------------------<DECLARACION DE VARIABLES>----------------------
            StreamWriter ParameterFile;
            //------------------------------------------------------------------
            if (PathFile == "" || PathFile == null)
                {
                return false;
                }
            try
                {
                ParameterFile = FileDiskOpenCreate(PathFile);
                for (int i = 0; i < Parameter.Length; i++)
                    {
                    FileDiskWrite(ParameterFile, Parameter[i]);
                    }
                FileDiskCloseWriter(ParameterFile);
                }
            catch
                {
                return false; // Error en Creación del Fichero
                }
            return true; // Creación del Fichero OK
            }

        //==============================================================================================
        //                                  Método StorageSize
        //==============================================================================================
        //FORMATO:
        //public string StorageSize(double Unidad)

        //DESCRIPCION:
        //Este método analiza la Unidad de almacenamiento expresa
        //y retorna una subcadena expresada en la unidad más conv
        //almacenamiento equivalente a Unidad.

        //PARAMETROS DE ENTRADA:

        //1) double Unidad
        //Contiene el tamaño del fichero expresado en bytes.
        //RETORNO:
        //Unidad >= 0 & Unidad < 1000
        //retorna Equivalente + " Bytes" --> Ejemplo "567 Bytes"
        //Unidad >= 1000 & Unidad < 1000000
        //retorna Equivalente + " Kilobytes" --> Ejemplo "567 Kb"
        //Unidad >= 1000000 & Unidad < 1000000000
        //retorna Equivalente + " Megabytes" --> Ejemplo "567 Mb"

        public string StorageSize(double Unidad)
            {
            if (Unidad >= 0 & Unidad < 1000)
                {
                return Unidad.ToString() + " Bytes";
                }
            if (Unidad >= 1000 & Unidad < 1000000)
                {
                return (Unidad / 1000).ToString() + " Kb";
                }
            if (Unidad >= 1000000 & Unidad < 1000000000)
                {
                return (Unidad / 1000000).ToString() + " Mb";
                }
            if (Unidad >= 1000000000 & Unidad < 1000000000000)
                {
                return (Unidad / 1000000000).ToString() + " Gb";
                }
            if (Unidad >= 1000000000000 & Unidad < 1000000000000000)
                {
                return (Unidad / 1000000000000).ToString() + " Tb";
                }
            return "Capacidad Erronea";
            }
        }
    }
