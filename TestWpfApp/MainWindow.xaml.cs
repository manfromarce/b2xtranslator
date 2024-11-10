using System;
using System.IO;
using System.Text;
using System.Windows;
using b2xtranslator.OpenXmlLib.WordprocessingML;
using b2xtranslator.OpenXmlLib.SpreadsheetML;
using b2xtranslator.OpenXmlLib.PresentationML;
using b2xtranslator.DocFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.PptFileFormat;
using b2xtranslator.StructuredStorage.Reader;
using Microsoft.Win32;
using Environment = System.Environment;

namespace TestWpfApp;
/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }

    private void Button_Click(object sender, RoutedEventArgs e)
    {
        var ofd = new OpenFileDialog()
        {
            Multiselect = true,
            Filter = "Office 97-2003 documents|*.doc;*.dot;*.xls;*.xlt;*.ppt;*.pps;*.pot",
        };
        if (ofd.ShowDialog(this) == true)
        {
            var folderDlg = new OpenFolderDialog()
            {
                Multiselect = false,
            };
            if (folderDlg.ShowDialog(this) == true)
            {
                string outputDir = folderDlg.FolderName;
                foreach (string file in ofd.FileNames)
                {
                    string inputExt = Path.GetExtension(file).ToLower();
                    try
                    {
                        using (var reader = new StructuredStorageReader(file))
                        {
                            string outputExt = inputExt + "x";
                            string baseName = Path.GetFileNameWithoutExtension(file);
                            string outputFile = Path.Join(outputDir, baseName + outputExt);
                            var outputType = b2xtranslator.OpenXmlLib.OpenXmlDocumentType.Document;
                            if (inputExt == ".dot" || inputExt == ".xlt" || inputExt == ".pot")
                            {
                                outputType = b2xtranslator.OpenXmlLib.OpenXmlDocumentType.Template;
                            }
                            switch (inputExt)
                            {
                                case ".doc":
                                case ".dot":
                                    var doc = new WordDocument(reader);
                                    using (var docx = WordprocessingDocument.Create(outputFile, outputType))
                                    {
                                        b2xtranslator.WordprocessingMLMapping.Converter.Convert(doc, docx);
                                    }
                                    break;
                                case ".xls":
                                case ".xlt":
                                    var xls = new XlsDocument(reader);
                                    using (var xlsx = SpreadsheetDocument.Create(outputFile, outputType))
                                    {
                                        b2xtranslator.SpreadsheetMLMapping.Converter.Convert(xls, xlsx);
                                    }
                                    break;
                                case ".ppt":
                                case ".pps":
                                case ".pot":
                                    var ppt = new PowerpointDocument(reader);
                                    using (var pptx = PresentationDocument.Create(outputFile, outputType))
                                    {
                                        b2xtranslator.PresentationMLMapping.Converter.Convert(ppt, pptx);
                                    }
                                    break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw;
                        //MessageBox.Show("Conversion failed: " + Environment.NewLine + ex.Message);
                    }
                }
            }
        }
    }
}