using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data;
using System.Xml;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace Civ6LanguageXmlConverter
{
    internal class Program
    {


        static string DirectoryPath;
        //static string ProjectFilePath;

        static DataTable TableSheet = new DataTable();
        static bool isUI = false;
        static List<string> Log = new List<string>();

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

                ErrorMessage("Error Args:You Need Input A Directory Path!");

            }
            else
            {
                OpenDialog();
            }

            Convert();
            OutputXls();
            LastDialog();

        }

        static void LastDialog()
        {

            if (!isUI) { return; }


            Form last = new Form();

            last.Width = 800;
            last.Height = 600;

            RichTextBox box = new RichTextBox();

            box.ReadOnly = true;

            foreach(string text in Log)
            {

                box.AppendText(text+'\n');

            }
            last.Text = "Log";
            last.MaximizeBox = false;
            last.MinimizeBox = false;
            box.Dock = DockStyle.Fill;
            last.Controls.Add(box);
            

            last.ShowDialog();
            Application.Exit();

        }

        static void BuildDataTable()
        {
            InfoMessage("Build Table...");
            TableSheet.Columns.Add("File");
            TableSheet.Columns.Add("Comment");
            TableSheet.Columns.Add("Tag");
            TableSheet.Columns.Add("Language");
            TableSheet.Columns.Add("Text");





        }

        
        static void AddLanguageToDataTable(string language)
        {
            InfoMessage("Add New Language:"+language);

            TableSheet.Columns.Add(language);

        }

        static void Convert()
        {
            BuildDataTable();
            InfoMessage("Try Convert:"+DirectoryPath);

            string[] xmls = Directory.GetFiles(DirectoryPath, "*.xml", SearchOption.AllDirectories);



            foreach(string file in xmls)
            {

                InfoMessage("Try get language from " + file);

                XmlDocument xml = new XmlDocument();

                try
                {

                    using (StreamReader reader = new StreamReader(file))
                    {
                        string code = reader.ReadToEnd();
                        string finalcode = Regex.Replace(code, "[\x00-\x08\x0B\x0C\x0E-\x1F\x26]", "", RegexOptions.Compiled);
                        xml.LoadXml(finalcode);


                        XmlNodeList list = xml.GetElementsByTagName("BaseGameText");
                        InfoMessage("Found " + list.Count + " <BaseGameText>!");

                        for(int i = 0; i < list.Count; i++)
                        {
                            AddGameText(Path.GetFileName(file), list.Item(0).ChildNodes);
                        }

                        
                        list = xml.GetElementsByTagName("LocalizedText");
                        InfoMessage("Found " + list.Count + " <LocalizedText>!");
                        for (int i = 0; i < list.Count; i++)
                        {
                            AddGameText(Path.GetFileName(file), list.Item(0).ChildNodes);
                        }
                    }
                    
                }
                catch (Exception e)
                {

                    DebugMessage(e.Message+"\n"+e.StackTrace);


                }
                finally
                {
                    
                }

                

            }


            ////<UpdateText>
            //InfoMessage("Search modinfo...");
            //string[] modinfos = Directory.GetFiles(DirectoryPath, "*.modinfo", SearchOption.TopDirectoryOnly);
            //InfoMessage("Search civ6proj...");
            //string[] civ6projs = Directory.GetFiles(DirectoryPath, "*.civ6proj", SearchOption.TopDirectoryOnly);

            //if (civ6projs.Length >= 0)
            //{
            //    InfoMessage("Get first civ6proj file:" + civ6projs[0]);
            //    InfoMessage("Pass modinfo beacause civb6proj file exists!");
            //}
            //else if (modinfos.Length >= 0 && civ6projs.Length <= 0)
            //{
            //    InfoMessage("Get first modinfo file:" + modinfos[0]);
            //}
            //else
            //{
            //    InfoMessage("Cannot get modinfo and civ6proj, Try search all xml files:" + DirectoryPath);
            //}




        }


        static void OutputXls()
        {
            string filename = Path.GetFileName(DirectoryPath);
            InfoMessage("Save xlsx:" + System.AppDomain.CurrentDomain.BaseDirectory + filename + ".xlsx");
            XLWorkbook book = new XLWorkbook();
            IXLWorksheet sheet =  book.Worksheets.Add(TableSheet, "Civ6LanguageXmlConverter_V000");
            sheet.Cells().Style.Alignment.SetWrapText(true);
            sheet.Cells().Style.Alignment.SetVertical(XLAlignmentVerticalValues.Top);
            sheet.Cells().Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);

            sheet.Column(1).Width = 48;
            sheet.Column(2).Width = 64;
            sheet.Column(3).Width = 48;
            sheet.Column(4).Width = 16;
            sheet.Column(5).Width = 128;

            book.SaveAs(System.AppDomain.CurrentDomain.BaseDirectory + filename+".xlsx");
            


        }


        static void AddGameText(string file,XmlNodeList list)
        {

            InfoMessage("Add game text from " + file + " ,Number:" + list.Count);

            for(int i = 0; i < list.Count; i++)
            {

                string comment="";
                string tag = "";
                string language = "";
                string text = "";

                if (list[i].Name == "Row")
                {

                    tag = list[i].Attributes["Tag"].Value;

                    if (i - 1 >= 0 && list[i].NodeType == XmlNodeType.Comment)
                    {
                        comment = list[i].InnerText;
                    }


                    if (list[i].Attributes["Language"] != null)
                    {
                        language = list[i].Attributes["Language"].Value;
                    }

                    if (list[i].HasChildNodes)
                    {
                        text = list[i].FirstChild.InnerText;
                    }


                    InfoMessage("Try add " + tag + " from " + file + " where language=" + language + " and text=" + text + " and comment=" + comment);
                    if (!String.IsNullOrEmpty(tag))
                    {
                        AddToTable(file,tag,language,text,comment);
                    }
                   


                }
                



            }
            

        }


        static bool AddToTable(string file,string tag,string language,string text,string comment)
        {
            

            DataRow row = TableSheet.NewRow();

            row["Tag"] = tag;
            row["File"] = file;
            row["Comment"] = comment;

            if (String.IsNullOrEmpty(language))
            {
                row["Language"] = "<BaseGameText>";
            }
            else
            {

                row["Language"] = language;

                //bool flag = true;
                //for(int i = 0; i < TableSheet.Columns.Count; i++)
                //{
                //    if (TableSheet.Columns[i].ColumnName == language.ToLower())
                //    {
                //        flag = false;
                //        break;
                //    }
                //}


                //if (flag)
                //{
                //    AddLanguageToDataTable(language.ToLower());
                //}

                //row[language.ToLower()] = text;
            }


            row["Text"] = text;

            TableSheet.Rows.Add(row);

            return true;
        }


        static bool VaildDirectory(string path)
        {


            if (Directory.Exists(path))
            {
                return true;
            }
            else
            {
                return false;
            }


        }


        static void ErrorMessage(string msg)
        {
            Log.Add("[ERROR]"+msg);
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
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "Select target directory:";

            if(dialog.ShowDialog() == DialogResult.OK)
            {

                DirectoryPath = dialog.SelectedPath;

            }
            else
            {


                MessageBox.Show("Error Args:You Need Input A Directory Path!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, false);
                ErrorMessage("Error Args:You Need Input A Directory Path!");
            }

        }



    }







}
