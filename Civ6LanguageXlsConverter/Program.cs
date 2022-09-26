using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace Civ6LanguageXlsConverter
{
    internal class Program
    {


        static string DirectoryPath;
        //static string ProjectFilePath;

        static bool isUI = false;
        static List<string> Log = new List<string>();
        static Dictionary<string, List<XmlFileData>> DataMap = new Dictionary<string, List<XmlFileData>>();


        struct XmlFileData
        {
            public string File;
            public string Language;
            public string Text;
            public string Comment;
            public string Tag;
            public bool Unknown;
        }


        

        [STAThread]
        static void Main(string[] args)
        {
            InfoMessage("Author:HaoJun0823:https://blog.haojun0823.xyz/");
            InfoMessage("Project:Github:https://github.com/HaoJun0823/Civ6LanguageConverter");
            InfoMessage("Used dependencies:Open-XML-SDK / Copyright (c) .NET Foundation and Contributors / MIT License / https://github.com/OfficeDev/Open-XML-SDK");
            InfoMessage("Used dependencies:ExcelNumberFormat / Copyright (c) 2017 andersnm / MIT License / https://github.com/andersnm/ExcelNumberFormat");
            InfoMessage("Used dependencies:ClosedXML / Copyright (c) 2016 ClosedXML / MIT License / https://github.com/ClosedXML/ClosedXML");
            InfoMessage("Used dependencies:Costura.Fody / Copyright (c) 2012 Simon Cropp and contributors / MIT License / https://github.com/Fody/Costura");
            InfoMessage("Used dependencies:Fody / Copyright (c) Simon Cropp and contributors / MIT License / https://github.com/Fody/Home/");

            if (args.Length != 0)
            {

                if (VaildDirectory(args[0]))
                {
                    DirectoryPath = args[0];
                }

                ErrorMessage("Error Args:You Need Input A Excel File Path!");

            }
            else
            {
                OpenDialog();
            }

            Convert();
            OutputXml();
            LastDialog();

        }




        static void Convert()
        {

            using(XLWorkbook excel = new XLWorkbook(DirectoryPath))
            {

                IXLWorksheet sheet = excel.Worksheet(1);
                int i = 2;

                int FileNumber = -1, CommentNumber = -1,TagNumber = -1,LanguageNumber = -1,TextNumber = -1;

                try
                {
                    InfoMessage("Search Columns Metadata On Row:" + i);
                    FileNumber = sheet.Row(1).Search("File").First<IXLCell>().WorksheetColumn().ColumnNumber();
                    InfoMessage("Get File Data On Column:" + FileNumber);
                    CommentNumber = sheet.Row(1).Search("Comment").First<IXLCell>().WorksheetColumn().ColumnNumber();
                    InfoMessage("Get Comment Data On Column:" + CommentNumber);
                    TagNumber = sheet.Row(1).Search("Tag").First<IXLCell>().WorksheetColumn().ColumnNumber();
                    InfoMessage("Get Tag Data On Column:" + TagNumber);
                    LanguageNumber = sheet.Row(1).Search("Language").First<IXLCell>().WorksheetColumn().ColumnNumber();
                    InfoMessage("Get Language Data On Column:" + LanguageNumber);
                    TextNumber = sheet.Row(1).Search("Text").First<IXLCell>().WorksheetColumn().ColumnNumber(); ;
                    InfoMessage("Get Text Data On Column:" + TextNumber);
                }
                catch(Exception e)
                {
                    
                    DebugMessage(e.Message);
                    DebugMessage(e.StackTrace);
                    ErrorMessage("Error Get Column!!!");
                    return ;
                }

                

                InfoMessage("Redirect On Row:" + i);
                while (!sheet.Row(i).IsEmpty())
                {
                    InfoMessage("Get Row " + i + " DATA:" + sheet.Row(i).ToString());
                    string File = sheet.Row(i).Cell(1).Value.ToString();
                    string Comment = sheet.Row(i).Cell(2).Value.ToString();
                    string Tag = sheet.Row(i).Cell(3).Value.ToString();
                    string Language = sheet.Row(i).Cell(4).Value.ToString();
                    string Text = sheet.Row(i).Cell(5).Value.ToString();


                    InfoMessage("File=" + File + ",Comment=" + Comment + "Tag=" + Tag + "Language=" + Language + "Text=" + Text);

                    XmlFileData data = new XmlFileData();
                    data.Text = Text;
                    data.File = File;
                    data.Comment = Comment;
                    data.Tag = Tag;
                    data.Language = Language;
                    

                    if(String.IsNullOrEmpty(Tag)|| String.IsNullOrEmpty(Text))
                    {
                        DebugMessage("Error Data Beacause Tag=Null Or Text=Null");
                        DebugMessage("Try Convert To Comment...");
                        data.Unknown = true;
                    }


                    if (!DataMap.ContainsKey(File))
                    {
                        InfoMessage("This Is A New File,Create List...");
                        DataMap.Add(File, new List<XmlFileData>());
                    }

                    DataMap[File].Add(data);


                    InfoMessage("Add Data To:" + File);
                    i++;
                }


            }

        }

        static void OutputXml()
        {

            string folder = System.AppDomain.CurrentDomain.BaseDirectory + Path.GetFileNameWithoutExtension(DirectoryPath)+".Language";

            InfoMessage("Lock Output Folder:"+folder);
            Directory.CreateDirectory(folder);

            foreach(var item in DataMap)
            {

                InfoMessage("Output Xml:" + item.Key + ",Count:" + item.Value.Count);
                XmlDocument xml = new XmlDocument();
                xml.AppendChild(xml.CreateXmlDeclaration("1.0", "utf-8", null));
                xml.AppendChild(xml.CreateComment(item.Key));
                xml.AppendChild(xml.CreateComment("Automatic generated by Civ6LanguageConverter"));
                xml.AppendChild(xml.CreateComment("Program Author:HaoJun0823:https://blog.haojun0823.xyz/"));
                xml.AppendChild(xml.CreateComment("Date Created:"+DateTime.Now));
                XmlElement GameData = xml.CreateElement("GameData");
                XmlElement BaseGameText = xml.CreateElement("BaseGameText");
                XmlElement LocalizedText = xml.CreateElement("LocalizedText");

                //GameData.AppendChild(GameBaseText);
                //GameData.AppendChild(LocalizedText);

                


                foreach (var data in item.Value)
                {
                    InfoMessage("Get Data,File=" + data.File + ",Comment=" + data.Comment + ",Tag=" + data.Tag + ",Language=" + data.Language + ",Text=" + data.Text + ",Unknown=" + data.Unknown);

                    if (data.Unknown)
                    {
                        InfoMessage("Unknown Data,Convert to Comment...");
                        GameData.AppendChild(xml.CreateComment("File=" + data.File + ",Comment=" + data.Comment + ",Tag=" + data.Tag + ",Language=" + data.Language + ",Text=" + data.Text + ",Unknown=" + data.Unknown));
                        continue;
                    }



                    if (data.Language.Equals("<BaseGameText>",StringComparison.OrdinalIgnoreCase))
                    {
                        InfoMessage("Insert Into BaseGameText...");
                        if (!String.IsNullOrEmpty(data.Comment))
                        {
                            InfoMessage("Insert Comment...");
                            BaseGameText.AppendChild(xml.CreateComment(data.Comment));


                        }
                        XmlElement element = xml.CreateElement("Row");
                        element.SetAttribute("Tag", data.Tag);
                        XmlElement xtext = xml.CreateElement("Text");
                        xtext.InnerText = data.Text;
                        element.AppendChild(xtext);
                        BaseGameText.AppendChild(element);


                    }
                    else
                    {
                        InfoMessage("Insert Into LocalizedText...");
                        if (!String.IsNullOrEmpty(data.Comment))
                        {
                            InfoMessage("Insert Comment...");
                            LocalizedText.AppendChild(xml.CreateComment(data.Comment));
                        }
                        XmlElement element = xml.CreateElement("Row");
                        element.SetAttribute("Tag", data.Tag);
                        element.SetAttribute("Language", data.Language) ;
                        XmlElement xtext = xml.CreateElement("Text");
                        xtext.InnerText = data.Text;
                        element.AppendChild(xtext);
                        LocalizedText.AppendChild(element);

                    }
                
                
                }




                if (LocalizedText.ChildNodes.Count <= 0)
                {
                    InfoMessage("Remove LocalizedText Because it doesn't have any child.");


                }
                else
                {
                    GameData.PrependChild(LocalizedText);
                }

                if (BaseGameText.ChildNodes.Count <= 0)
                {
                    InfoMessage("Remove BaseGameText Because it doesn't have any child.");


                }
                else
                {
                    GameData.PrependChild(BaseGameText);
                }

                xml.AppendChild(GameData);


                xml.Save(folder +"//" +item.Key);
                InfoMessage("Done:" + folder+"//"+item.Key);

            }


            



        }

        static void LastDialog()
        {

            if (!isUI) { return; }


            Form last = new Form();

            last.Width = 800;
            last.Height = 600;

            RichTextBox box = new RichTextBox();

            box.ReadOnly = true;

            foreach (string text in Log)
            {

                box.AppendText(text + '\n');

            }
            last.Text = "Log";
            last.MaximizeBox = false;
            last.MinimizeBox = false;
            box.Dock = DockStyle.Fill;
            last.Controls.Add(box);


            last.ShowDialog();
            Application.Exit();

        }





        static void ErrorMessage(string msg)
        {
            Log.Add("[ERROR]" + msg);
            Console.Error.WriteLine("[ERROR]{0}", msg);
            LastDialog();
        }

        static void InfoMessage(string msg)
        {
            Log.Add("[INFO]" + msg);
            Console.WriteLine("[INFO]{0}", msg);
        }

        static void DebugMessage(string msg)
        {
            Log.Add("[DEBUG]" + msg);
            Console.WriteLine("[DEBUG]{0}", msg);
        }



        static void OpenDialog()
        {

            isUI = true;
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Select target excel file:";
            dialog.Filter = "Excel Worksheets|*.xls;*.xlsx";

            if (dialog.ShowDialog() == DialogResult.OK)
            {

                DirectoryPath = dialog.FileName;

            }
            else
            {


                MessageBox.Show("Error Args:You Need Input A Excel File Path!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, false);
                ErrorMessage("Error Args:You Need Input A Excel File Path!");
            }

        }


        static bool VaildDirectory(string path)
        {


            if (File.Exists(path))
            {
                return true;
            }
            else
            {
                return false;
            }


        }




    }
}
