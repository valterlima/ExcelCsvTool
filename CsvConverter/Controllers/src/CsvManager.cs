using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Web;
using System.IO.Compression;

namespace CsvConverter.Controllers.src
{
    public class CsvManager
    {
        private HttpPostedFileBase file;
        private string fileName;
        private DataSet result;
        private List<string> csvFiles;
        private DirectoryInfo tempDirectory;
        public string zipFilePath { get; set; }

        public CsvManager(HttpPostedFileBase file)
        {
            this.file = file;
            this.ReadFile();
            this.GenerateCsvFiles();
            this.zipAllFiles();
        }

        private void ReadFile()
        {
            if (file.ContentLength <= 0)
            {
                throw new Exception("File is empty.");
            }
            else
            {
                MemoryStream ms = new MemoryStream();
                file.InputStream.CopyTo(ms);

                IExcelDataReader reader = null;
                try
                {
                    string type = file.ContentType;
                    this.fileName = file.FileName.Replace(".xlsx", "").Replace(".xls", "");
                    //application/vnd.ms-excel
                    //application/vnd.openxmlformats-officedocument.spreadsheetml.sheet

                    if (type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    {
                        reader = ExcelReaderFactory.CreateOpenXmlReader(ms);
                    }
                    else if (type == "application/vnd.ms-excel")
                    {
                        reader = ExcelReaderFactory.CreateBinaryReader(ms);
                    }
                    else
                    {
                        throw new Exception("Invalid format: " + type);
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("ReaderFactory error", ex);
                }
                finally
                {
                    this.result = reader.AsDataSet();
                    reader.Close();
                }
            }
        }

        private void GenerateCsvFiles()
        {
            List<string> csvFiles = new List<string>();
            string fileName = this.fileName;
            string directoryName = DateTime.UtcNow.ToString("yyyy-MM-ddTHH-mm-ssZ");
            var saveDirPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/UploadedFiles"), directoryName);
            this.tempDirectory = new DirectoryInfo(saveDirPath);

            Directory.CreateDirectory(saveDirPath);

            foreach (DataTable sheet in result.Tables)
            {
                string sheetName = sheet.TableName.ToString();
                string csvString = "";

                foreach (DataRow row in sheet.Rows)
                {
                    foreach (Object col in row.ItemArray)
                    {
                        csvString = String.Concat(csvString, col.ToString(), ",");
                    }
                    csvString = String.Concat(csvString, "\n");
                }

                string newFileName = String.Concat(fileName, "-", sheetName, ".csv");
                string saveFilePath = Path.Combine(saveDirPath, newFileName);

                StreamWriter csv = new StreamWriter(saveFilePath, false);
                csv.Write(csvString);
                csv.Close();

                csvFiles.Add(saveFilePath);
            }

            this.csvFiles = csvFiles;
        }

        private void zipAllFiles()
        {
            string zipFilePath = Path.Combine(this.tempDirectory.Parent.FullName, this.tempDirectory.Name + "." + this.fileName + ".zip");

            try
            {
                ZipFile.CreateFromDirectory(this.tempDirectory.FullName, zipFilePath);

                this.zipFilePath = zipFilePath;
            }
            catch(Exception ex)
            {
                throw new Exception("Error while trying to save files.", ex);
            }
        }
    }
}