using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows;
using ExcelDataReader;
using System.Windows.Input;
using System.Windows.Controls;

namespace Missing
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        public static RoutedCommand RunRoutedCommand = new RoutedCommand();
        public static RoutedCommand OpenFileRoutedCommand = new RoutedCommand();

        enum MType { FirstName, Surname, Mobile, Email, Status, LicenceType, ValidTill };

        IList<IDictionary<MType, string>> mCIMembers = new List<IDictionary<MType, string>>();
        IList<IDictionary<MType, string>> mTextMembers = new List<IDictionary<MType, string>>();
        IList<IDictionary<MType, string>> mEmailMembers = new List<IDictionary<MType, string>>();

        const string mLogFile = @"reports.txt";

        String[] GetFiles(string title, string filter)
        {
            Microsoft.Win32.OpenFileDialog FileOpen = new Microsoft.Win32.OpenFileDialog();

            FileOpen.Title = title;
            FileOpen.Filter = filter;
            FileOpen.CheckFileExists = true;
            FileOpen.CheckPathExists = true;
            FileOpen.Multiselect = true;

            if (FileOpen.ShowDialog() == true)
            {
                return FileOpen.FileNames;
            }
            else
            {
                return null;
            }
        }

        private void CanOpenFileCommand(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
            e.Handled = true;
        }

        private void OpenFileCommand(object sender, ExecutedRoutedEventArgs e)
        {
            e.Handled = true;
            string Title = string.Empty;
            string Filter = string.Empty;

            switch(e.Parameter)
            {
                case "Email":
                    Title = "Select an email export comma separated file";
                    Filter = "CSV files (*.csv)|*.csv|(*.txt)|*.txt|All files (*.*)|*.*";
                    break;
                case "Text":
                    Title = "Select a mobile numbers comma seperated file";
                    Filter = "CSV files (*.csv)|*.csv|(*.txt)|*.txt|All files (*.*)|*.*";
                    break;
                case "CI":
                    Title = "Select an excel Cycling Ireland members file";
                    Filter = "CSV files (*.xls)|*.xls|All files (*.*)|*.*";
                    break;
                default:
                    return;
            }


            string[] Files = GetFiles(Title, Filter);

            if ( Files != null && Files.Length == 1)
            {
                switch(e.Parameter)
                {
                    case "Email":
                        mEmailList.Text = Files[0];                        
                        break;
                    case "Text":
                        mTextNumbers.Text = Files[0];
                        break;
                    case "CI":
                        mCyclingIreland.Text = Files[0];
                        break;
                    default:
                        return;
                }
            }
        }

        private void RunCommand(object sender, ExecutedRoutedEventArgs e)
        {
            try
            {
                
                if (File.Exists(mLogFile))
                    File.Delete(mLogFile);
            }
            catch
            {
                MessageBox.Show(string.Format("Failed to delete old log file {0}", mLogFile), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            mListBox.Items.Clear();

            FindMissingTextNumbers();

            FindNonMembersInTextNumbers();

            FindMissingEmailAddresses();

            FindNotRegisteredThisYear();

        }

        private void CanRunCommand(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = (mCIMembers.Any() && mTextMembers.Any() && mEmailMembers.Any());
            e.Handled = true;
        }

        public MainWindow()
        {
            InitializeComponent();

            mCyclingIreland.PreviewDragOver += (s, e) => 
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files != null && files.Length == 1)
                {
                    if ( File.Exists(files[0]))
                    {
                        if (IsCSVFile(files[0]) == false && IsExcel(files[0]))
                            (s as TextBox).Text = files[0];
                    }                    
                }
                e.Handled = true;
            };

            mCyclingIreland.TextChanged += (s, e) =>
            {
                string FileName = (s as TextBox).Text;
                if (File.Exists(FileName) && IsExcel(FileName) )
                    mCIMembers = ReadCyclingIrelandMembers(FileName);
            };

            mTextNumbers.PreviewDragOver += (s, e) =>
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files != null && files.Length == 1)
                {
                    if (File.Exists(files[0]))
                    {
                        if (IsCSVFile(files[0]))
                            (s as TextBox).Text = files[0];
                    }
                }
                e.Handled = true;
            };

            mTextNumbers.TextChanged += (s, e) =>
            {
                string FileName = (s as TextBox).Text;
                if (File.Exists(FileName) && IsCSVFile(FileName))
                    mTextMembers = ReadTextMembers(FileName);
            };

            mEmailList.PreviewDragOver += (s, e) =>
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files != null && files.Length == 1)
                {
                    if (File.Exists(files[0]))
                    {
                        if (IsCSVFile(files[0]))
                            (s as TextBox).Text = files[0];
                    }
                }
                e.Handled = true;
            };
            mEmailList.TextChanged += (s, e) =>
            {
                string FileName = (s as TextBox).Text;
                if (File.Exists(FileName) && IsCSVFile(FileName))
                    mEmailMembers = ReadEmailMembers(FileName);
            };

            //mCIMembers = ReadCyclingIrelandMembers(@"C:\Users\Ger\Desktop\IMBRC\clubmembers.xls");
            //mTextMembers = ReadTextMembers(@"C:\Users\Ger\Desktop\IMBRC\text.numbers.txt");
            //mEmailMembers = ReadEmailMembers(@"C:\Users\Ger\Desktop\IMBRC\mail.chimp.export.csv");

            CommandBinding CustomCommandBinding = new CommandBinding(RunRoutedCommand, RunCommand, CanRunCommand);
            this.CommandBindings.Add(CustomCommandBinding);
            mRun.Command = RunRoutedCommand;

            CustomCommandBinding = new CommandBinding(OpenFileRoutedCommand, OpenFileCommand, CanOpenFileCommand);
            this.CommandBindings.Add(CustomCommandBinding);

            mEmailFileOpen.Command = OpenFileRoutedCommand;
            mCIFileOpen.Command = OpenFileRoutedCommand;
            mTextFileOpen.Command = OpenFileRoutedCommand;


            mCyclingIreland.Text = @"C:\Users\Ger\Desktop\IMBRC\clubmembers.xls";
            mEmailList.Text = @"C:\Users\Ger\Desktop\IMBRC\mail.chimp.export.csv";
            mTextNumbers.Text = @"C:\Users\Ger\Desktop\IMBRC\exportedcontacts.txt";

        }

        

        void Log(string filename, string text)
        {
            if (string.IsNullOrEmpty(filename) || string.IsNullOrEmpty(text))
                return;
            try
            {
                StreamWriter File = new StreamWriter(filename, true);
                File.WriteLine(text);
                File.Close();
            }
            catch
            {

            }
        }

        void Log(string filename, string[] lines)
        {
            if (string.IsNullOrEmpty(filename) || lines == null || lines.Length == 0)
                return;
            try
            {
                StreamWriter File = new StreamWriter(filename, true);
                foreach ( string Line in lines)
                    File.WriteLine(Line);
                File.Close();
            }
            catch
            {

            }
        }

        private void FindNotRegisteredThisYear()
        {
            mListBox.Items.Add("Members not registered");
            IList<IDictionary<MType, string>> Members = new List<IDictionary<MType, string>>();
            int Missing = 0;
            foreach (var Member in mCIMembers.Where(x => x[MType.Status] != "Registered")  )
            {
                if (string.IsNullOrEmpty(Member[MType.ValidTill]))
                    continue;
                DateTime ValidUntil = DateTime.Now;
                if ( DateTime.TryParse(Member[MType.ValidTill], out ValidUntil) )
                {
                    DateTime _2016 = new DateTime(2016, 1, 1);
                    if (DateTime.Compare(ValidUntil, _2016) > 0)
                    {
                        Missing++;
                        mListBox.Items.Add(string.Format("{0} {1} {2} [{3}] - {4}", Member[MType.FirstName], Member[MType.Surname], Member[MType.Email], Member[MType.LicenceType], ValidUntil.ToShortDateString()));
                        Members.Add(Member);
                    }
                }
            }
            if (Missing == 0)
            {
                mListBox.Items.Add("No unregistered members");
            }
            else
            {
                mListBox.Items.Add(string.Format("{0} members not registered", Missing));

                mListBox.Items.Add(string.Format("Non Senior and Junior members", Missing));
                if (Members.Any(x => x[MType.LicenceType].StartsWith("U")))
                {
                    foreach (IDictionary<MType, string> Member in Members.Where(x => x[MType.LicenceType].StartsWith("U")))
                    {
                        string Text = string.Format("{0}, {1} {2},{3},{4}", Member[MType.Email], Member[MType.FirstName], Member[MType.Surname], Member[MType.LicenceType], Member[MType.ValidTill]);
                        Log(mLogFile, Text);
                        mListBox.Items.Add(Text);
                    }
                }
                mListBox.Items.Add(string.Format("Senior and Junior members", Missing));
                if (Members.Any(x => x[MType.LicenceType].StartsWith("U") == false))
                {
                    foreach (IDictionary<MType, string> Member in Members.Where(x => x[MType.LicenceType].StartsWith("U") == false))
                    {
                        string Text = string.Format("{0}, {1} {2},{3},{4}", Member[MType.Email], Member[MType.FirstName], Member[MType.Surname], Member[MType.LicenceType], Member[MType.ValidTill]);
                        Log(mLogFile, Text);
                        mListBox.Items.Add(Text);

                    }
                }
            }
        }

        private void FindMissingEmailAddresses()
        {
            int Missing=0;
            mListBox.Items.Add("Registered members not in Email list");
            IList<IDictionary<MType, string>> Members = new List<IDictionary<MType, string>>();
            Missing = 0;
            foreach (var Member in mCIMembers.Where(x => x[MType.Status] == "Registered" && string.IsNullOrEmpty(x[MType.LicenceType]) == false && x[MType.LicenceType].StartsWith("U") == false))
            {
                if (mEmailMembers.Any(x => x[MType.Email] == Member[MType.Email]) == false)
                {
                    Missing++;
                    Members.Add(Member);
                    mListBox.Items.Add(string.Format("{0} {1} [ {2} ] {3}", Member[MType.FirstName], Member[MType.Surname], Member[MType.Email], Member[MType.LicenceType]));
                }
            }

            if (Missing == 0)
            {
                mListBox.Items.Add("No missing members");
            }
            else
            {
                mListBox.Items.Add(string.Format("{0} members not found in email list", Missing));
                Log(mLogFile, string.Format("{0} members not found in email list", Missing));
                foreach (IDictionary<MType, string> Member in Members)
                    Log(mLogFile, string.Format("{0}, {1} {2},{3}", Member[MType.Email], Member[MType.FirstName], Member[MType.Surname], Member[MType.LicenceType]));
            }
        }

        private void FindMissingTextNumbers()
        {
            int Missing = 0;
            mListBox.Items.Add("Registered members not in Text list");
            IList<IDictionary<MType, string>> Members = new List<IDictionary<MType, string>>();
            foreach (var Member in mCIMembers.Where(x => x[MType.Status] == "Registered" && string.IsNullOrEmpty(x[MType.LicenceType]) == false && x[MType.LicenceType].StartsWith("U") == false))
            {
                if (mTextMembers.Any(x => x[MType.Mobile] == Member[MType.Mobile]) == false)
                {
                    Missing++;
                    string Text = string.Format("{0} {1} [ {2} ] {3}", Member[MType.FirstName], Member[MType.Surname], Member[MType.Mobile], Member[MType.LicenceType]);
                    Members.Add(Member);
                    mListBox.Items.Add(Text);
                }
            }

            if (Missing == 0)
            {
                mListBox.Items.Add("No missing members");
            }
            else
            {
                mListBox.Items.Add(string.Format("{0} members not found in text list", Missing));
                Log(mLogFile, string.Format("{0} members not found in text list", Missing));
                foreach (IDictionary<MType, string> Member in Members)
                    Log(mLogFile, string.Format("{0}, {1} {2},{3}", Member[MType.Mobile], Member[MType.FirstName], Member[MType.Surname], Member[MType.LicenceType]));
            }
            
        }

        private void FindNonMembersInTextNumbers()
        {
            int Missing = 0;
            mListBox.Items.Add("Non registered members in Text list");
            IList<IDictionary<MType, string>> Members = new List<IDictionary<MType, string>>();
            foreach (var Member in mCIMembers.Where(x => x[MType.Status] != "Registered" && string.IsNullOrEmpty(x[MType.LicenceType]) == false && x[MType.LicenceType].StartsWith("U") == false))
            {
                if (mTextMembers.Any(x => x[MType.Mobile] == Member[MType.Mobile]))
                {
                    Missing++;
                    string Text = string.Format("{0} {1} [ {2} ] {3}", Member[MType.FirstName], Member[MType.Surname], Member[MType.Mobile], Member[MType.LicenceType]);
                    Members.Add(Member);
                    mListBox.Items.Add(Text);
                }
            }

            if (Missing == 0)
            {
                mListBox.Items.Add("No non registered members in Text list");
            }
            else
            {
                mListBox.Items.Add(string.Format("{0} non registered members found in text list", Missing));
                Log(mLogFile, string.Format("{0} non registered members found in text list", Missing));
                foreach (IDictionary<MType, string> Member in Members)
                    Log(mLogFile, string.Format("{0}, {1} {2},{3}", Member[MType.Mobile], Member[MType.FirstName], Member[MType.Surname], Member[MType.LicenceType]));
            }

        }

        private IList<IDictionary<MType, string>> ReadTextMembers(string filename)
        {
            IList<IDictionary<MType, string>> TextMembers = new List<IDictionary<MType, string>>();
            if (IsCSVFile(filename))
            {
                foreach ( string Line in File.ReadAllLines(filename))
                {
                    IDictionary<MType, string> MemberData = new Dictionary<MType, string>();
                    string Normalised = Line.Replace("\"", "").Replace("+", "");

                    string[] Tokens = Normalised.Split(',');

                    double Number = 0;
                    int n = 0;
                    for (n = 0; n < Tokens.Length; n++)
                    {
                        if (double.TryParse(Tokens[n], out Number))
                            break;
                    }
                    if ( Number > 0 )
                    {                        
                        String FullName = string.Empty;
                        if (n > 0)
                            FullName = Tokens[0];
                        else if (Tokens.Length > 1)
                            FullName = Tokens[1];

                        if (string.IsNullOrEmpty(FullName) == false)
                        {
                            MemberData.Add(MType.Mobile, NormaliseNumber(Number.ToString()));
                            MemberData.Add(MType.FirstName, GetFirstname(FullName));
                            MemberData.Add(MType.Surname, GetSurname(FullName));
                            TextMembers.Add(MemberData);
                        }
                    }
                }
            }
            return TextMembers;
        }

        private string NormaliseNumber(string number)
        {
            if (number.StartsWith("353"))
                return number.Substring(3);
            return number;
        }

        private string GetSurname(string fullName)
        {
            int Pos = fullName.IndexOf(" ");
            if ( Pos > 0 )
            {
                return fullName.Substring(Pos, fullName.Length-Pos);
            }
            return fullName;
        }

        private string GetFirstname(string fullName)
        {
            int Pos = fullName.IndexOf(" ");
            if (Pos > 0)
            {
                return fullName.Substring(0, Pos);
            }
            return fullName;
        }

        private IList<IDictionary<MType, string>> ReadEmailMembers(string filename)
        {
            IList<IDictionary<MType, string>> EmailMembers = new List<IDictionary<MType, string>>();
            if (IsCSVFile(filename))
            {
                string[] Lines = File.ReadAllLines(filename);

                string[] Tokens = Lines[0].Replace("\"", "").Split(',');

                IDictionary<MType, int> ColIndex = new Dictionary<MType, int>();

                for ( int i=0 ; i<Tokens.Length ; i++ )
                {
                    switch(Tokens[i])
                    {
                        case "Email Address":
                            ColIndex.Add(MType.Email, i);
                            break;
                        case "First Name":
                            ColIndex.Add(MType.FirstName, i);
                            break;
                        case "Last Name":
                            ColIndex.Add(MType.Surname, i);
                            break;
                        case "Mobile Number":
                            ColIndex.Add(MType.Mobile, i);
                            break;
                    }
                }

                if (ColIndex.Any() == false)
                    return EmailMembers;

                int Max = ColIndex.Values.Max();

                foreach (string Line in Lines)
                {
                    IDictionary<MType, string> MemberData = new Dictionary<MType, string>();
                    Tokens = Line.Split(',');
                    if (Tokens.Length > Max)
                    {
                        foreach (KeyValuePair<MType, int> KV in ColIndex)
                        {
                            String Value = Tokens[KV.Value];
                            switch (KV.Key)
                            {
                                case MType.Mobile:
                                    double Number = 0;
                                    if (double.TryParse(Value, out Number))
                                        Value = Number.ToString();
                                    break;
                                default:
                                        Value = Value.ToLowerInvariant();
                                    break;
                            }
                            MemberData.Add(KV.Key, Value);
                        }

                        EmailMembers.Add(MemberData);
                    }
                }
            }
            return EmailMembers;
        }

        private IList<IDictionary<MType, string>> ReadCyclingIrelandMembers(string filename)
        {
            IList<IDictionary<MType, string>> CIMembers = new List<IDictionary<MType, string>>();
            System.Data.DataTable XcelMembers = ReadExcel(filename);
            if (XcelMembers != null && XcelMembers.Rows != null && XcelMembers.Rows.Count > 1)
            {
                IDictionary<MType, int> ColumnIndex = new Dictionary<MType, int>();
                DataRow ColNames = XcelMembers.Rows[0];

                for (int i = 0; i < ColNames.ItemArray.Length; i++)
                {
                    switch (ColNames[i].ToString())
                    {
                        case "Email Address":
                            ColumnIndex.Add(MType.Email, i);
                            break;
                        case "Mobile Number":
                            ColumnIndex.Add(MType.Mobile, i);
                            break;
                        case "CI Status":
                            ColumnIndex.Add(MType.Status, i);
                            break;
                        case "Firstname":
                            ColumnIndex.Add(MType.FirstName, i);
                            break;
                        case "Surname":
                            ColumnIndex.Add(MType.Surname, i);
                            break;
                        case "Licence Type":
                            ColumnIndex.Add(MType.LicenceType, i);
                            break;
                        case "Valid Till":
                            ColumnIndex.Add(MType.ValidTill, i);
                            break;
                            /*
                                                    case "Licence Type":
                                                    case "Category":
                                                    case "Licence Number":
                                                    case "Valid Till":
                                                    */
                    }
                }

                for (int n = 1; n < XcelMembers.Rows.Count; n++)
                {
                    DataRow RowData = XcelMembers.Rows[n];
                    IDictionary<MType, string> MemberData = new Dictionary<MType, string>();

                    foreach (KeyValuePair<MType, int> KV in ColumnIndex)
                    {
                        String Data = RowData[KV.Value].ToString().Trim();
                        switch (KV.Key)
                        {
                            case MType.FirstName:
                            case MType.Surname:
                            case MType.Email:
                                Data = Data.ToLowerInvariant();
                                break;

                            case MType.Mobile:
                                Data = Data.Replace("+", "").Replace("(", "").Replace(")", "").Replace(" ", "");
                                if (Data.StartsWith("00"))
                                    Data = Data.Substring(2);
                                if (Data.StartsWith("353") && Data.Length > 7)
                                    Data = Data.Substring(3);
                                if (Data.StartsWith("00"))
                                    Data = Data.Substring(2);

                                double Number = 0;

                                if (double.TryParse(Data, out Number))
                                    Data = Number.ToString();

                                break;
                        }
                        MemberData.Add(KV.Key, Data);
                    }

                    CIMembers.Add(MemberData);

                }                
            }
            return CIMembers;
        }

        private bool IsCSVFile(string filePath)
        {
            if (File.Exists(filePath) == false)
                return false;
            string[] Lines = File.ReadAllLines(filePath);
            if (Lines.Length < 1)
                return false;
            return Lines[0].Split(',').Length > 1;
        }

        bool IsExcel(string fileName)
        {
            return (ReadExcel(fileName) != null);
        }

        System.Data.DataTable ReadExcel(string fileName)
        {            
            using (var stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx)
                using (IExcelDataReader Reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = Reader.AsDataSet();

                    if (result.Tables != null && result.Tables.Count > 0)
                    {
                        foreach (var Table in result.Tables )
                        {
                            if (Table.ToString() == @"Club Members + Recent Club Fee")
                                return Table as System.Data.DataTable;
                        }
                    }
                    // The result of each spreadsheet is in result.Tables
                }
                return null;
            }
        }
    }
}
