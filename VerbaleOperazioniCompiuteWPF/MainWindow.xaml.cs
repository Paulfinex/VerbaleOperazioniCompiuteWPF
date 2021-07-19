using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Word=Microsoft.Office.Interop.Word;
using System.ComponentModel;

namespace VerbaleOperazioniCompiuteWPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {

        bool debug = true;
        static string RunningPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
        string savePath = "";
        List<string> lista_comuni = new List<string>(new string[] {"Firenze","Pistoia","Prato","Perugia","Lucca","Massa","Pisa","Arezzo","Siena","Bari","Ancona","Termini Imerese","Massa","Livorno","La Spezia" });
        List<string> lista_comuni2 = new List<string>(new string[] { "Firenze", "Pistoia", "Prato", "Perugia", "Lucca", "Massa", "Pisa", "Arezzo", "Siena", "Bari", "Ancona", "Termini Imerese", "Massa", "Livorno", "La Spezia" });
        List<string> lista_consulenti = new List<string>(new string[]{"Venerino Lo Cicero", "Andrea Bigagli","Gianluca Di Maria", "Francesco Pini"});
        List<string> lista_tipo_processo = new List<string>(new string[] { "R.G.N.R" });
        List<string> lista_mod_processo = new List<string>(new string[] { "Mod.21", "Mod.44", "Mod.52" });
        List<string> lista_tipo_rep = new List<string>(new string[] { "PC", "NoteBook", "Cellulare", "Tablet", "Lettore MP3", "Lettore MP4", "Hard Disk", "CD", "Pen Drive USB", "Fotocamera", "Videocamera", "Micro SD", "Posta", "Social",  "SD", "Floppy", "Hard Disk USB", "Router", "SIM Card", "NAS" });
        List<string> lista_marca_rep = new List<string>(new string[] { "SAMSUNG", "ASUS", "APPLE", "XIAOMI", "HUAWEI", "ONE PLUS", "NOKIA", "ACER", "MOTOROLA", "MICROSOFT", "OPPO", "KINGSTON", "APACER", "WIKO", "HP", "TRASCEND", "WD", "LENOVO", "BQ AQUARIS", "SEAGATE", "L8 STAR", "ZTE", "VODAFONE", "WINDOWS PHONE", "SANDISK", "PACKARD BELL", "SONY", "PANASONIC", "FUJITSU", "ARCHOS", "ALCATEL", "MSI", "RAZER", "DELL", "MICROSOFT", "MAXTOR", "IBM", "TOSHIBA", "BRONDI", "PNY", "VERBATIM", "NIKON", "GOOGLE", "GMAIL", "FACEBOOK", "INSTAGRAM", "OUTLOOK", "ONE DRIVE", "LG", "MEDIACOM" });
        List<string> lista_social = new List<string>(new string[] { "Gmail","Instagram", "Facebook", "TikTok", "WeChat", "Whatsapp", "iTunes", "Libero","Hotmail", "Alice", "Telegram", "Aruba", "Dropbox" });
        List<Reperto> lista_reperti = new List<Reperto>();
        List<Indagato> lista_indagati = new List<Indagato>();
        List<Avvocato> lista_avvocati = new List<Avvocato>();
        private readonly BackgroundWorker worker = new BackgroundWorker();
        string baseOperazioniCompiute = "";
        string baseAttivitaPeritali="";
        int counterIndagati = 1;
        public MainWindow()
        {
            InitializeComponent();
            if (debug)
            {
                baseAttivitaPeritali = string.Format("{0}Resources\\baseAttivitaPeritali.docx", System.IO.Path.GetFullPath(System.IO.Path.Combine(RunningPath, @"..\..\")));
                baseOperazioniCompiute = string.Format("{0}Resources\\baseOperazioniCompiute.docx", System.IO.Path.GetFullPath(System.IO.Path.Combine(RunningPath, @"..\..\")));

            }
            else
            {
                baseAttivitaPeritali = System.IO.Path.GetDirectoryName(RunningPath) + @"\Resources\baseAttivitaPeritali.docx"; 
                baseOperazioniCompiute = System.IO.Path.GetDirectoryName(RunningPath) + @"\Resources\baseOperazioniCompiute.docx"; 

            };
            MessageBox.Show("Seleziona una cartella di lavoro.");
            savePath = openFolderDialog();

            worker.DoWork += worker_DoWork;
            worker.RunWorkerCompleted += worker_RunWorkerCompleted;
            lbl_data_inizio_op.Content = DateTime.Now.ToString("dd'/'MM'/'yyyy");
            lbl_ora_inizio_op.Content = DateTime.Now.ToString("HH':'mm");
            combo_luogo_operazione.ItemsSource = lista_comuni;
            combo_luogo_tribunale.ItemsSource = lista_comuni2;
            combo_modello_procedimento.ItemsSource = lista_mod_processo;
            combo_tipo_procedimento.ItemsSource = lista_tipo_processo;
            combo_consulente.ItemsSource = lista_consulenti;
            txt_tipo_reperto.ItemsSource = lista_tipo_rep;
            txt_marca_reperto.ItemsSource = lista_marca_rep;

            lbl_nome_avvocato.Visibility = Visibility.Hidden;
            txt_nome_avvocato.Visibility = Visibility.Hidden;
            lbl_foro.Visibility = Visibility.Hidden;
            txt_foro_avvocato.Visibility = Visibility.Hidden;
            lbl_indagato.Visibility = Visibility.Hidden;
            txt_indagato_avvocato.Visibility = Visibility.Hidden;
            lbl_ctp.Visibility = Visibility.Hidden;
            txt_CTP.Visibility = Visibility.Hidden;
            listBox_CTP.Visibility = Visibility.Hidden;
            btn_add_CTP.Visibility = Visibility.Hidden;
            btn_rem_CTP.Visibility = Visibility.Hidden;
            listBox_avvocato.Visibility = Visibility.Hidden;
            btn_add_avvocato.Visibility = Visibility.Hidden;
            btn_rem_avvocato.Visibility = Visibility.Hidden;
            btn_export_account.Visibility = Visibility.Hidden;


        }

        private void btn_add_ausiliario_Click(object sender, RoutedEventArgs e)
        {
            if (isFilled(txt_ausiliario.Text))
            {
                listBox_ausiliario.Items.Add(txt_ausiliario.Text);
                txt_ausiliario.Text = "";
            }  
        }

        private void btn_add_indagato_Click(object sender, RoutedEventArgs e)
        {
            if (isFilled(txt_indagato.Text))
            {
                listBox_indagato.Items.Add(txt_indagato.Text);
                Indagato ind = new Indagato(txt_indagato.Text, counterIndagati++);
                lista_indagati.Add(ind);
                
                txt_indagato.Text = "";

            }
        }

        private bool isFilled(string s)
        {
            if (s != "")
                return true;
            return false;
        
        
        }

        private void btn_rem_ausiliario_Click(object sender, RoutedEventArgs e)
        {
            listBox_ausiliario.Items.Remove(listBox_ausiliario.SelectedItem);
        }

        private void btn_rem_indagato_Click(object sender, RoutedEventArgs e)
        {
            listBox_indagato.Items.Remove(listBox_indagato.SelectedItem);
            foreach (Indagato i in lista_indagati)
            {
                if (i.nome == listBox_indagato.SelectedItem)
                {
                    lista_indagati.Remove(i);
                }
            
            }
        }

        private string openDialog()
        {
            Microsoft.Win32.OpenFileDialog wordDialog = new Microsoft.Win32.OpenFileDialog();
            wordDialog.DefaultExt = ".doc";
            Nullable<bool> result = wordDialog.ShowDialog();
            if (result == true)
            {
                return wordDialog.FileName;
            }
            return null;

        }
        private string openFolderDialog() 
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                System.Windows.Forms.DialogResult result = dialog.ShowDialog();
                return dialog.SelectedPath;
            }

            return null;
        }

        private void ReplaceBookmarkText(Microsoft.Office.Interop.Word.Document doc, string bookmarkName, string text)

        {
            try
            {
                if (doc.Bookmarks.Exists(bookmarkName))

                {

                    Object name = bookmarkName;

                    Microsoft.Office.Interop.Word.Range range = doc.Bookmarks.get_Item(ref name).Range;



                    range.Text = text;

                    object newRange = range;

                    doc.Bookmarks.Add(bookmarkName, ref newRange);

                }
            }
            catch (Exception saveexception)
            {
                MessageBox.Show("Errore bookmark, riprovare inserimento." + saveexception);
            }
        }
 

        private void btn_aggiungi_reperto_Click(object sender, RoutedEventArgs e)
        {
            if (listBox_indagato.SelectedIndex == -1)
            {
                MessageBox.Show("Selezionare un indagato dalla lista indagati.");
                return;
            }

            int temp = 0;
            int currentRepId = 0;
            foreach (Indagato i in lista_indagati)
            {
                if (i.nome == (string)listBox_indagato.SelectedItem)
                {
                    temp = i.id;
                    i.counterReperti++;
                    currentRepId = i.counterReperti;


                }
            }
            Reperto r = new Reperto("R_"+ txt_num_procedimento.Text+"_"+ temp.ToString() + "_" + currentRepId.ToString(), txt_tipo_reperto.Text, txt_marca_reperto.Text, txt_modello_reperto.Text, txt_IMEI_reperto.Text, listBox_indagato.SelectedItem.ToString(),txt_password_rep.Text,txt_PIN_rep.Text,txt_condizioni_rep.Text);
            if (MessageBox.Show("Inserire il seguente reperto?" + Environment.NewLine + Environment.NewLine + r.ToString(), "Inserimento Reperto", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.No)
            {
                return;
            }
            else
            {
                lista_reperti.Add(r);
                listBox_reperti.ItemsSource = null;
                listBox_reperti.ItemsSource = lista_reperti;
                resetFormReperto();
                txt_IMEI_reperto.Text = "";
                txt_modello_reperto.Text = "";
                txt_tipo_reperto.Text = "";
                txt_marca_reperto.Text = "";
                txt_password_rep.Text = "Non Fornito";
                txt_PIN_rep.Text = "Non Fornito";
                txt_condizioni_rep.Text = "Normali condizioni d'uso";
            }

            
        }

        public class Avvocato 
        {
            public string nome { get; set; }
            public string foro { get; set; }
            public string indagati { get; set; }

            public Avvocato(string nome, string foro, string indagati)
            {
                this.nome = nome;
                this.foro = foro;
                this.indagati = indagati;

            }

            public override string ToString()
            {
                return "Nome: " + nome + " - Del foro di: " + foro + " - Indagati: " + indagati;
            }


        }
        public class Reperto
        {


            public string num_reperto { get; set; }
            public string tipo_reperto { get; set; }
            public string marca_reperto { get; set; }
            public string modello_reperto { get; set; }
            public string imei__reperto { get; set; }
            public string proprietario__reperto { get; set; }

            public string password_reperto { get; set; }
            public string pin_reperto { get; set; }
            public string condizioni_reperto { get; set; }
            public int sottoRepCount { get; set; }
            public Reperto(string num_reperto, string tipo_reperto, string marca_reperto, string modello_reperto, string imei__reperto, string proprietario__reperto, 
                string password_reperto,string pin_reperto, string condizioni_reperto)
            {
                this.num_reperto = num_reperto;
                this.tipo_reperto = tipo_reperto;
                this.marca_reperto = marca_reperto;
                this.modello_reperto = modello_reperto;
                this.imei__reperto = imei__reperto;
                this.proprietario__reperto = proprietario__reperto;
                this.password_reperto = password_reperto;
                this.pin_reperto = pin_reperto;
                this.condizioni_reperto = condizioni_reperto;
                this.sottoRepCount = 0;
            }
            public override string ToString()
            {
                if (tipo_reperto.Equals("account", StringComparison.InvariantCultureIgnoreCase))
                {
                    return "Numero Reperto:" + num_reperto + "\n" + "In uso a:" + proprietario__reperto + "\n" + "Tipo:" + tipo_reperto + "\n" + "Piattaforma:" + marca_reperto + "\n" +
                    "Username:" + modello_reperto + "\n" +
                    "Password:" + password_reperto + "\n";
                }
                else 
                {
                    return "Numero Reperto:" + num_reperto + "\n" + "In uso a:" + proprietario__reperto + "\n" + "Tipo:" + tipo_reperto + "\n" +
                        "Marca:" + marca_reperto + "\n" + "Modello:" + modello_reperto + "\n";
                }

                

            }
          
        }
            public class Indagato
            {


                public string nome { get; set; }
                public int id { get; set; }
            public int counterReperti { get; set; }
                public string proprietario__reperto { get; set; }
                public Indagato(string nome, int id)
                {
                    this.nome = nome;
                    this.id = id;
                this.counterReperti = 0;

                }

         
            
    
             }

        private void btn_inserisci_Click(object sender, RoutedEventArgs e)
        {

            if (txt_num_procedimento.Text == "")
            {
                MessageBox.Show("E' necessario inserire il numero di procedimento.");

            
            }

            string path = "";
            lista_reperti.Sort((p, q) => p.num_reperto.CompareTo(q.num_reperto));
            if (radio_verbale1.IsChecked == true)
            {
                path = baseOperazioniCompiute;
            }
            else
            {
                path = baseAttivitaPeritali;
            }

            ArrayList parole = new ArrayList();
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            object fileName = path;
            object confirmConversions = Type.Missing;
            object readOnly = Type.Missing;
            object addToRecentFiles = Type.Missing;
            object passwordDoc = Type.Missing;
            object passwordTemplate = Type.Missing;
            object revert = Type.Missing;
            object writepwdoc = Type.Missing;
            object writepwTemplate = Type.Missing;
            object format = Type.Missing;
            object encoding = Type.Missing;
            object visible = Type.Missing;
            object openRepair = Type.Missing;
            object docDirection = Type.Missing;
            object notEncoding = Type.Missing;
            object xmlTransform = Type.Missing;
            Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Open(
            ref fileName, ref confirmConversions, ref readOnly, ref addToRecentFiles,
            ref passwordDoc, ref passwordTemplate, ref revert, ref writepwdoc,
            ref writepwTemplate, ref format, ref encoding, ref visible, ref openRepair,
            ref docDirection, ref notEncoding, ref xmlTransform);

            try
            {
                if (radio_verbale1.IsChecked == true)
                {
                    doc.SaveAs(savePath + "\\VerbaleOperazioniCompiute" + txt_num_procedimento.Text + ".docx");
                }
                else
                {
                    doc.SaveAs(savePath + "\\VerbaleAttivitaPeritali" + txt_num_procedimento.Text + ".docx");
                }
            }
            catch (Exception e2)
            {
                MessageBox.Show("Errore nel salvataggio del File doc: " + e2.Message);
                
            }
            try
            {
                if (radio_verbale1.IsChecked == true)
                {
                    ReplaceBookmarkText(doc, "num_procedimento", txt_num_procedimento.Text);
                    ReplaceBookmarkText(doc, "anno_procedimento", txt_anno_procedimento.Text);
                    ReplaceBookmarkText(doc, "tipo_procedimento", combo_tipo_procedimento.Text);
                    ReplaceBookmarkText(doc, "modello_procedimento", combo_modello_procedimento.Text);
                    ReplaceBookmarkText(doc, "luogo_operazione", combo_luogo_operazione.Text);
                    ReplaceBookmarkText(doc, "data_operazione", lbl_data_inizio_op.Content.ToString());
                    ReplaceBookmarkText(doc, "tribunale", combo_luogo_tribunale.Text);
                    ReplaceBookmarkText(doc, "tribunale2", combo_luogo_tribunale.Text);
                    string gender = "del Dr. ";

                    if (radioF.IsChecked == true)
                    {
                        gender = "della Dr.ssa ";
                    }

                    if (radioPM.IsChecked == true)
                    {
                        if (radioF.IsChecked == true)
                        {
                            gender = " Dr.ssa ";
                        }
                        else
                        {
                            gender = " Dr. ";
                        }
                    }
                    ReplaceBookmarkText(doc, "pm", gender + txt_pm.Text);
                    ReplaceBookmarkText(doc, "pm2", gender + txt_pm.Text);
                    ReplaceBookmarkText(doc, "sottoscritto", "Il sottoscritto ");
                    ReplaceBookmarkText(doc, "consulente", combo_consulente.Text);
                    string pg = "Ausiliario di P.G. dalla " + txt_pg.Text + " su delega";
                    if (radioPM.IsChecked == true)
                    {
                        pg = "dal ";
                    }
                    ReplaceBookmarkText(doc, "pg", pg);
                    ReplaceBookmarkText(doc, "data_delega", data_delega.Text);
                    string indagati = "";
                    foreach (string s in listBox_indagato.Items)
                    {
                        indagati += s + " e ";
                    }
                    ReplaceBookmarkText(doc, "indagati", indagati.Substring(0, indagati.Length - 3));

                    ReplaceBookmarkText(doc, "data_inizio_op", lbl_data_inizio_op.Content.ToString());
                    ReplaceBookmarkText(doc, "ora_inizio_op", lbl_ora_inizio_op.Content.ToString());

                    int countind = 0;
                    string indagatiBuild = "";
                    foreach (string s in listBox_indagato.Items)
                    {
                        countind++;
                    }
                    if (countind > 1)
                    {
                        indagatiBuild += "dagli indagati ";
                        foreach (string s in listBox_indagato.Items)
                        {
                            indagatiBuild += s + ", ";
                        }

                    }
                    else
                    {
                        indagatiBuild += "dall' indagato ";
                        foreach (string s in listBox_indagato.Items)
                        {
                            indagatiBuild += s;
                        }
                    }
                    ReplaceBookmarkText(doc, "indagato_individua", indagatiBuild);
                    ReplaceBookmarkText(doc, "luogo_attività", txt_box_attivita.Text);
                    ReplaceBookmarkText(doc, "note_indagato", txt_note.Text);

                    string ausiliari = "";
                    string templateAusiliari = "Per le attività peritali odierne mi sono avvalso";
                    int count = 0;
                    foreach (string s in listBox_ausiliario.Items)
                    {
                        ausiliari = ausiliari + s + " e ";
                        count++;
                    }
                    if (count > 1)
                    {
                        templateAusiliari += " dei miei collaboratori " + ausiliari;
                    }
                    else
                    {
                        templateAusiliari += " del mio collaboratore " + ausiliari;
                    }
                    if (count > 0)
                        ReplaceBookmarkText(doc, "fine_ausiliari", templateAusiliari.Substring(0, templateAusiliari.Length - 3));
                    ReplaceBookmarkText(doc, "ora_fine_op", DateTime.Now.ToString("HH':'mm"));
                    ReplaceBookmarkText(doc, "data_fine_op", DateTime.Now.ToString("dd'/'MM'/'yyyy"));
                    ReplaceBookmarkText(doc, "num_copie", txt_copie.Text);
                    if (count > 0)
                        ReplaceBookmarkText(doc, "ausiliari_chiusura", ausiliari.Substring(0, ausiliari.Length - 3));
                    ReplaceBookmarkText(doc, "consulente_chiusura", combo_consulente.Text);
                    ReplaceBookmarkText(doc, "dichiarazioni", txt_dichiarazioni.Text);
                }
                if (radio_verbale2.IsChecked == true)
                {
                    ReplaceBookmarkText(doc, "data_corrente", lbl_data_inizio_op.Content.ToString());
                    ReplaceBookmarkText(doc, "luogo_attivita", combo_luogo_operazione.Text);
                    ReplaceBookmarkText(doc, "data_conferimento", data_delega.Text);
                    ReplaceBookmarkText(doc, "num_procedimento", txt_num_procedimento.Text);
                    ReplaceBookmarkText(doc, "data_procedimento", txt_anno_procedimento.Text);
                    ReplaceBookmarkText(doc, "mod_procedimento", combo_modello_procedimento.Text);
                    string gender = "del Dr. ";

                    if (radioF.IsChecked == true)
                    {
                        gender = "della Dr.ssa ";
                    }

                    if (radioPM.IsChecked == true)
                    {
                        if (radioF.IsChecked == true)
                        {
                            gender = " Dr.ssa ";
                        }
                        else
                        {
                            gender = " Dr. ";
                        }
                    }
                    ReplaceBookmarkText(doc, "PM", gender + txt_pm.Text);
                    ReplaceBookmarkText(doc, "tribunale", combo_luogo_tribunale.Text);

                    int countavvocati = 0;
                    foreach (string i in listBox_avvocato.Items)
                    {
                        countavvocati++;
                    }
                    if (countavvocati > 0)
                    {

                        ReplaceBookmarkText(doc, "gliavvocat", "gli avvocati");
                    }
                    else
                    {
                        ReplaceBookmarkText(doc, "gliavvocat", "l'avvocato");
                    }
                    string builderAvvocati = "";

                    foreach (Avvocato a in lista_avvocati)
                    {
                        builderAvvocati += a.nome + " del foro di " + a.foro + ", difensore degli indagati " + a.indagati;
                        builderAvvocati += " - ";

                    }
                    ReplaceBookmarkText(doc, "avvocati_foro", builderAvvocati);
                    string ctp_builder = "";
                    foreach (string ctp in listBox_CTP.Items)
                    {
                        ctp_builder += ctp + " - ";

                    }

                    ReplaceBookmarkText(doc, "CTP", ctp_builder);

                    string ausiliaribuilder = "";
                    foreach (string a in listBox_ausiliario.Items)
                    {
                        ausiliaribuilder += a + " , ";
                    }
                    ReplaceBookmarkText(doc, "ausiliari", ausiliaribuilder);

                    ReplaceBookmarkText(doc, "num_proc_piepagina", txt_num_procedimento.Text);
                    ReplaceBookmarkText(doc, "data_proc_piepagina", txt_anno_procedimento.Text);
                    ReplaceBookmarkText(doc, "tribunale_piepagina", combo_luogo_tribunale.Text);
                    ReplaceBookmarkText(doc, "pm_piepagina", txt_pm.Text);

                    ReplaceBookmarkText(doc, "ora_conclusione2", DateTime.Now.ToString("HH':'mm"));
                    ReplaceBookmarkText(doc, "data_conclusione2", DateTime.Now.ToString("dd'/'MM'/'yyyy"));
                    ReplaceBookmarkText(doc, "CTP2", ctp_builder);
                    ReplaceBookmarkText(doc, "CTP3", ctp_builder);
                    ReplaceBookmarkText(doc, "num_copie", txt_copie.Text);
                    ReplaceBookmarkText(doc, "consulente", combo_consulente.Text);
                    ReplaceBookmarkText(doc, "ausiliari2", ausiliaribuilder);
                    ReplaceBookmarkText(doc, "ora_inizio_attivita", lbl_ora_inizio_op.Content.ToString());

                }
                if (doc.Bookmarks.Exists("tabelle"))
                {


                    Object name = "tabelle";
                    Object start = doc.Bookmarks.get_Item(ref name).Start - 1;
                    Object end = doc.Bookmarks.get_Item(ref name).End;
                    Word.Range rng = doc.Range(start, start);
                    foreach (Reperto r in lista_reperti)
                    {



                        Word.Table table = doc.Tables.Add(rng, 3, 9);
                        rng = table.Range;
                        rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        table.Columns.DistributeWidth();
                        table.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                        table.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                        table.AllowAutoFit = true;
                        table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);

                        setCell(table, 1, 1, "Reperto N.");
                        setCell(table, 1, 2, "Tipo");
                        setCell(table, 1, 3, "Marca");
                        setCell(table, 1, 4, "Modello");
                        setCell(table, 1, 5, "IMEI");
                        setCell(table, 1, 6, "In uso a");
                        setCell(table, 1, 7, "Codice/Password");
                        setCell(table, 1, 8, "PIN SIM");
                        setCell(table, 1, 9, "Condizioni");
                        table.Cell(1, 1).Range.Shading.BackgroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlueGray;
                        table.Cell(1, 2).Range.Shading.BackgroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlueGray;
                        table.Cell(1, 3).Range.Shading.BackgroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlueGray;
                        table.Cell(1, 4).Range.Shading.BackgroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlueGray;
                        table.Cell(1, 5).Range.Shading.BackgroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlueGray;
                        table.Cell(1, 6).Range.Shading.BackgroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlueGray;
                        table.Cell(1, 7).Range.Shading.BackgroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlueGray;
                        table.Cell(1, 8).Range.Shading.BackgroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlueGray;
                        table.Cell(1, 9).Range.Shading.BackgroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlueGray;
                        setCell(table, 2, 1, r.num_reperto);
                        setCell(table, 2, 2, r.tipo_reperto);
                        setCell(table, 2, 3, r.marca_reperto);
                        setCell(table, 2, 4, r.modello_reperto);
                        setCell(table, 2, 5, r.imei__reperto);
                        setCell(table, 2, 6, r.proprietario__reperto);
                        setCell(table, 2, 7, r.password_reperto);
                        setCell(table, 2, 8, r.pin_reperto);
                        setCell(table, 2, 9, r.condizioni_reperto);

                        
                        
                        table.Cell(3, 1).Merge(table.Cell(3, 9));


                        rng.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                        rng.InsertParagraphAfter();
                        rng.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                    }

                }

                doc.Save();
                doc.Close();
                wordApp.Quit();
                MessageBox.Show("Il documento salvato in: " + Environment.NewLine + savePath );

            }
            catch (Exception Ex_Inserimento)
            {

                doc.Close();
                wordApp.Quit();
                MessageBox.Show("Errore inserimento:" + Ex_Inserimento.Message);

            }
        }
        private void setCell(Word.Table table,int i ,int j,string text)
        {

            table.Cell(i,j).Range.Text = text;
            table.Cell(i,j).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
        }
        private bool checkFormFill()
        {
            

            return true;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            foreach (string s in listBox_ausiliario.Items)
            {
                MessageBox.Show(s);
            
            }
        }

        private void btn_sottoreperto_Click(object sender, RoutedEventArgs e)
        {
            if (listBox_reperti.SelectedIndex == -1)
            {
                MessageBox.Show("Selezionare un elemento dalla lista reperti per aggiungere un sottoreperto.");
                return;
            }
            string[] tmp = listBox_reperti.SelectedItem.ToString().Split('\n');
            string[] tmp1 = tmp[0].Split(':');
            string num_r = tmp1[1];
            tmp1 = tmp[1].Split(':');
            string nome_ind = tmp1[1];
            num_r = num_r.Replace("\n", "");
            nome_ind = nome_ind.Replace("\n", "");
            int temp = 0;
            int currentRepId = 0;
            int currentSottRepId = 0;
            int count = num_r.Split('_').Length - 1;

            if (count > 3)
            {
                MessageBox.Show("Non puoi inserire un sottoreperto di un sottoreperto, cambiare la selezione nella lista.");
                return;
            
            }
            foreach (Indagato i in lista_indagati)
            {
                if (i.nome == nome_ind)
                {
                    temp = i.id;
                    currentRepId = i.counterReperti;

                }
            }
            foreach (Reperto r1 in lista_reperti)
            {

                if (r1.num_reperto == num_r)
                {
                    r1.sottoRepCount++;
                    currentSottRepId = r1.sottoRepCount;
                }

            }

            Reperto r = new Reperto("R_" + txt_num_procedimento.Text + "_" + temp.ToString() + "_" + currentRepId.ToString() +"_"+currentSottRepId, txt_tipo_reperto.Text, txt_marca_reperto.Text, txt_modello_reperto.Text, txt_IMEI_reperto.Text, listBox_indagato.SelectedItem.ToString(), txt_password_rep.Text, txt_PIN_rep.Text, txt_condizioni_rep.Text);
            if (MessageBox.Show("Inserire il seguente reperto?" + Environment.NewLine + Environment.NewLine + r.ToString(), "Inserimento Reperto", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.No)
            {
                return;
            }
            else
            {
                lista_reperti.Add(r);
                listBox_reperti.ItemsSource = null;
                listBox_reperti.ItemsSource = lista_reperti;
                resetFormReperto();
                txt_IMEI_reperto.Text = "";
                txt_modello_reperto.Text = "";
                txt_tipo_reperto.Text = "";
                txt_marca_reperto.Text = "";
                txt_password_rep.Text = "Non Fornito";
                txt_PIN_rep.Text = "Non Fornito";
                txt_condizioni_rep.Text = "Normali condizioni d'uso";

            }

        }

        private void btn_export_account_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Selezionare cartella di destinazione.");
            string path = openFolderDialog();
            List<Reperto> lista_account = new List<Reperto>();
            foreach (Reperto r in lista_reperti)
            {
                if (r.tipo_reperto.Equals("account", StringComparison.InvariantCultureIgnoreCase))
                {
                    lista_account.Add(r);
                }


            }
            string builder = "";
            foreach (Reperto r in lista_account)
            {
                builder += r.ToString();
                builder += "_____________________" + Environment.NewLine;
            
            }

            Crea_txt(builder, path, "account" + txt_num_procedimento.Text);
        }
        private void Crea_txt(string testo, string percorso, string nome)
        {
            string fileName = percorso + @"\" + nome + ".txt";

            try
            {
                if (File.Exists(fileName))
                {
                    File.Delete(fileName);
                }
                using (StreamWriter sw = File.CreateText(fileName))
                {
                    sw.WriteLine(testo);
                }
            }
            catch (Exception Ex)
            {
                Console.WriteLine(Ex.ToString());
            }
        }

        private void check_account_Checked(object sender, RoutedEventArgs e)
        {
            txt_tipo_reperto.Text = "Account";
            lbl_marca.Content = "Piattaforma";
            txt_marca_reperto.ItemsSource = lista_social;
            lbl_modello.Content = "Username";
            txt_IMEI_reperto.IsEnabled = false;
            txt_PIN_rep.Text = "";
            txt_PIN_rep.IsEnabled = false;
            txt_condizioni_rep.Text = "";
            txt_condizioni_rep.IsEnabled = false;
        }
        private void resetFormReperto()
        {
            check_account.IsChecked = false;
            txt_tipo_reperto.Text = "";
            lbl_marca.Content = "Marca";
            txt_marca_reperto.ItemsSource = lista_marca_rep;
            lbl_modello.Content = "Modello";
            txt_IMEI_reperto.IsEnabled = true;
            txt_PIN_rep.Text = "Non Fornito";
            txt_PIN_rep.IsEnabled = true;
            txt_condizioni_rep.Text = "Normali condizioni d'uso";
            txt_condizioni_rep.IsEnabled = true;

        }

        private void radio_verbale2_Checked(object sender, RoutedEventArgs e)
        {
            //Attività peritale
            lbl_pg.Visibility = Visibility.Hidden;
            txt_pg.Visibility = Visibility.Hidden;
            lbl_conferitoda.Visibility = Visibility.Hidden; ;
            radioPM.Visibility = Visibility.Hidden; 
            radioPG.Visibility = Visibility.Hidden;
            txt_note.Visibility = Visibility.Hidden;
            lbl_note.Visibility= Visibility.Hidden;
            lbl_datadecreto.Content = "Data Conferimento";

            lbl_nome_avvocato.Visibility = Visibility.Visible;
            txt_nome_avvocato.Visibility = Visibility.Visible;
            lbl_foro.Visibility = Visibility.Visible;
            txt_foro_avvocato.Visibility = Visibility.Visible;
            lbl_indagato.Visibility = Visibility.Visible;
            txt_indagato_avvocato.Visibility = Visibility.Visible;
            lbl_ctp.Visibility = Visibility.Visible;
            txt_CTP.Visibility = Visibility.Visible;
            listBox_CTP.Visibility = Visibility.Visible;
            btn_add_CTP.Visibility = Visibility.Visible;
            btn_rem_CTP.Visibility = Visibility.Visible;
            listBox_avvocato.Visibility = Visibility.Visible;
            btn_add_avvocato.Visibility = Visibility.Visible;
            btn_rem_avvocato.Visibility= Visibility.Visible;

            btn_export_account.Visibility = Visibility.Visible;
        }

    

        private void radio_verbale2_Unchecked(object sender, RoutedEventArgs e)
        {//Verbale Operazioni compiute
            lbl_pg.Visibility = Visibility.Visible;
            txt_pg.Visibility = Visibility.Visible;
            lbl_conferitoda.Visibility = Visibility.Visible;
            radioPM.Visibility = Visibility.Visible;
            radioPG.Visibility = Visibility.Visible;
            txt_note.Visibility = Visibility.Visible;
            lbl_note.Visibility = Visibility.Visible;
            lbl_datadecreto.Content = "Data Decreto";

            lbl_nome_avvocato.Visibility = Visibility.Hidden;
            txt_nome_avvocato.Visibility = Visibility.Hidden;
            lbl_foro.Visibility = Visibility.Hidden;
            txt_foro_avvocato.Visibility = Visibility.Hidden;
            lbl_indagato.Visibility = Visibility.Hidden;
            txt_indagato_avvocato.Visibility = Visibility.Hidden;
            lbl_ctp.Visibility = Visibility.Hidden;
            txt_CTP.Visibility = Visibility.Hidden;
            listBox_CTP.Visibility = Visibility.Hidden;
            btn_add_CTP.Visibility = Visibility.Hidden;
            btn_rem_CTP.Visibility = Visibility.Hidden;
            listBox_avvocato.Visibility = Visibility.Hidden;
            btn_add_avvocato.Visibility = Visibility.Hidden;
            btn_rem_avvocato.Visibility = Visibility.Hidden;

            btn_export_account.Visibility = Visibility.Hidden;
        }

        private void btn_add_CTP_Click(object sender, RoutedEventArgs e)
        {
            if (isFilled(txt_CTP.Text))
            {
                listBox_CTP.Items.Add(txt_CTP.Text);
                txt_CTP.Text = "";

            }
        }

        private void btn_add_avvocato_Click(object sender, RoutedEventArgs e)
        {
            if (isFilled(txt_nome_avvocato.Text) && isFilled(txt_foro_avvocato.Text) && isFilled(txt_indagato_avvocato.Text) )
            {
                listBox_avvocato.Items.Add(txt_nome_avvocato.Text);
                Avvocato avv = new Avvocato(txt_nome_avvocato.Text, txt_foro_avvocato.Text, txt_indagato_avvocato.Text);
                lista_avvocati.Add(avv);

                listBox_indagato.Items.Add(txt_indagato_avvocato.Text);

                Indagato ind = new Indagato(txt_indagato_avvocato.Text, counterIndagati++);
                lista_indagati.Add(ind);
                txt_nome_avvocato.Text = "";
                txt_foro_avvocato.Text = "";
                txt_indagato_avvocato.Text = "";

            }
        }

        private void btn_rem_avvocato_Click(object sender, RoutedEventArgs e)
        {
            listBox_avvocato.Items.Remove(listBox_avvocato.SelectedItem);
            foreach (Avvocato a in lista_avvocati)
            {
                if (a.nome == listBox_avvocato.SelectedItem)
                {
                    lista_avvocati.Remove(a);
                }

            }
        }

        private void btn_rem_CTP_Click(object sender, RoutedEventArgs e)
        {
            listBox_CTP.Items.Remove(listBox_CTP.SelectedItem);
        }

        private void check_account_Unchecked(object sender, RoutedEventArgs e)
        {
            txt_tipo_reperto.Text = "Tipo";
            lbl_marca.Content = "Marca";
            txt_marca_reperto.ItemsSource = lista_marca_rep;
            lbl_modello.Content = "Modello";
            txt_IMEI_reperto.IsEnabled = true;
            txt_PIN_rep.Text = "Non Fornito";
            txt_PIN_rep.IsEnabled = true;
            txt_condizioni_rep.Text = "Normali condizioni d'uso";
            txt_condizioni_rep.IsEnabled = true;
        }

        private void worker_DoWork(object sender, DoWorkEventArgs e)
        {
        }

        private void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            
        }

        
    }
}

