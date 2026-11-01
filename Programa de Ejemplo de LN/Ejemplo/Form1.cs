using LN;
namespace Ejemplo
    {
    public partial class Form1 : Form
        {
        int NumeroDeFicheros = 0;
        string[] ListaDeFicheros;
        StreamReader ENTRADA;
        LN.LN OBJ = new LN.LN();
        public Form1()
            {
            InitializeComponent();
            }
        private void button1_Click(object sender, EventArgs e)
            {
            string Articulo = "";
            ENTRADA = OBJ.FileDiskOpenRead(@"D:\VIN.txt");
            Articulo = OBJ.FileDiskRead(ENTRADA);
            OBJ.MensINFO("Articulo = " + Articulo);
            //------------------------------------------------
            ListaDeFicheros = OBJ.GetDirectoryFilesDialog(ref NumeroDeFicheros); 
            OBJ.MensINFO("Número de Ficheros = " + NumeroDeFicheros.ToString());
            //------------------------------------------------
            OBJ.MSDOSWait("D:\\AAA.docx");
            }
        }
    }
