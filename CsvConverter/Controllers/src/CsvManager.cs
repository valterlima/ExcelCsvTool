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
        private List<FileInfo> files;
        private DirectoryInfo tempDirectory;
        public string zipFilePath;
        public string zipFileName { get; set; }

        public CsvManager(HttpPostedFileBase file)
        {
            this.file = file;
            this.ReadFile();
            this.GenerateCsvFiles();
        }

        private void ReadFile()
        {
            if (this.file.ContentLength <= 0)
            {
                throw new Exception("File is empty.");
            }
            else if (file.ContentLength > 1 * 1024 * 1024)
            {
                throw new Exception("File is too large. (Bigger than 1 MB)");
            }
            else
            {
                MemoryStream ms = new MemoryStream();
                this.file.InputStream.CopyTo(ms);

                IExcelDataReader reader = null;
                try
                {
                    string type = this.file.ContentType;
                    this.fileName = this.file.FileName.Replace(".xlsx", "").Replace(".xls", "");
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
                        throw new Exception("Invalid format");
                    }

                    this.result = reader.AsDataSet();
                    reader.Close();
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }
            }
        }

        private void GenerateCsvFiles()
        {
            this.csvFiles = new List<string>();
            this.files = new List<FileInfo>();

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
                files.Add(new FileInfo(saveFilePath));
            }
        }

        public string GenerateDownloadPath()
        {
            if (this.csvFiles.Count > 1)
            {
                this.zipAllFiles();
                return this.zipFileName;
            }
            else if (this.csvFiles.Count == 1)
            {
                return Path.Combine(this.tempDirectory.Name, this.files[0].Name);
            }
            else
            {
                return "";
            }
        }

        private void zipAllFiles()
        {
            string zipFileName = this.tempDirectory.Name + "." + this.fileName + ".zip";
            string zipFilePath = Path.Combine(this.tempDirectory.Parent.FullName, zipFileName);

            try
            {
                ZipFile.CreateFromDirectory(this.tempDirectory.FullName, zipFilePath);

                this.zipFilePath = zipFilePath;
                this.zipFileName = zipFileName;
            }
            catch (Exception ex)
            {
                throw new Exception("Error while trying to save files.", ex);
            }
        }
    }
}