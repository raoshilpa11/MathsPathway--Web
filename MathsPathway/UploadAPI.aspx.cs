using System;
using System.IO;
using System.Web;
using System.Web.Script.Serialization;
using ASPSnippets.GoogleAPI;
using Microsoft.Office.Interop.Excel;

namespace MathsPathway
{
    public partial class UploadAPI : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            GoogleConnect.ClientId = "client_id";
            GoogleConnect.ClientSecret = "client_secret";
            GoogleConnect.RedirectUri = Request.Url.AbsoluteUri.Split('?')[0];
            GoogleConnect.API = EnumAPI.Drive;
            if (!string.IsNullOrEmpty(Request.QueryString["code"]))
            {
                string code = Request.QueryString["code"];
                string json = GoogleConnect.PostFile(code, (HttpPostedFile)Session["File"], Session["Description"].ToString());
                GoogleDrive file = (new JavaScriptSerializer()).Deserialize<GoogleDrive>(json);
                tblFileDetails.Visible = true;
                lblTitle.Text = file.Title;
                lblId.Text = file.Id;
                imgIcon.ImageUrl = file.IconLink;
                lblCreatedDate.Text = file.CreatedDate.ToString();
                lnkDownload.NavigateUrl = file.WebContentLink;
                if (!string.IsNullOrEmpty(file.ThumbnailLink))
                {
                    rowThumbnail.Visible = true;
                    imgThumbnail.ImageUrl = file.ThumbnailLink;
                }
            }
            if (Request.QueryString["error"] == "access_denied")
            {
                ClientScript.RegisterClientScriptBlock(this.GetType(), "alert", "alert('Access denied.')", true);
            }

        }

        protected void UploadFile(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                
            }


            xlApp.DisplayAlerts = false;
            string filePath = @"C:\Users\Shilpa\Downloads\" + ((System.Web.UI.WebControls.FileUpload)(FindControl("FileUpload1"))).FileName;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(filePath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Microsoft.Office.Interop.Excel.Sheets worksheets = xlWorkBook.Worksheets;

            Worksheet wkSheet = new Worksheet();
            xlApp.DisplayAlerts = false;
            //for (int i = xlApp.ActiveWorkbook.Worksheets.Count; i > 0; i--)
            for (int i = xlApp.ActiveWorkbook.Worksheets.Count; i > 0; i--)
            {
                //wkSheet = (Worksheet)xlApp.ActiveWorkbook.Worksheets[i];
                wkSheet = xlWorkBook.Worksheets[i];
                if (wkSheet.Name != "Table 2")
                {
                    wkSheet.Delete();
                }
            }
            xlApp.DisplayAlerts = true;            
            xlWorkBook.Save();

            string filename = Path.GetFileName(FileUpload1.FileName);
            FileUpload1.SaveAs(Server.MapPath("~/") + filename);
            //HttpPostedFile upload = Server.MapPath("~/") + filename;
            xlWorkBook.Close();

            //releaseObject(worksheets);wkSheet
            releaseObject(wkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            //Session["File"] = filePath;  
            HttpFileCollection UploadedFiles = Request.Files;
            HttpPostedFile UserPostedFile;
            int UploadFileCount = UploadedFiles.Count;
            if (UploadFileCount >= 1)
            {
                for (int i = 0; i < UploadFileCount; ++i)
                {
                    UserPostedFile = UploadedFiles[i];
                    UserPostedFile.SaveAs(@"C:\Users\Shilpa\Downloads\" + UserPostedFile.FileName);
                    Session["File"] = UserPostedFile;
                }
            }

            //Session["File"] = FileUpload1.PostedFile; //UserPostedFile.ToString(); //FileUpload1.PostedFile;
            Session["Description"] = txtDescription.Text;
            GoogleConnect.Authorize("https://www.googleapis.com/auth/drive.file");
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;                
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}