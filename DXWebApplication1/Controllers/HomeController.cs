using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DXWebApplication1.Models;
using DevExpress.Web.Mvc;
using DevExpress.Web.Office;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;

namespace DXWebApplication1.Controllers
{
    public class HomeController : Controller
    {
        //
        // GET: /Home/

        private readonly MongoDBService _mongoDBService;
        private readonly List<ProjectModel> _projects;
        private readonly List<CompModel> _comps;
        private static ObjectId selectedProject;

        public HomeController()
        {
            _mongoDBService = new MongoDBService();
            _projects = _mongoDBService.GetProjects();
            _comps = _mongoDBService.GetComps();
        }

        public ActionResult Index(string Id = null)
        {
            if (Request != null && Request.QueryString["Id"] != null)
            {
                string projectId = Request.QueryString["Id"];
                selectedProject = _projects.Where(x => x.Id.ToString() == projectId).FirstOrDefault().Id;
            }
            else
            {
                selectedProject = _projects.FirstOrDefault().Id;
            }
            return View(_projects);
        }

        [HttpPost, ValidateInput(false)]
        public void IndexCallBack([ModelBinder(typeof(DevExpressEditorsBinder))] ProjectOptionModel options)
        {
            Console.WriteLine(options.ProjectName);
        }

        public ActionResult RichEditPartial(string actioName = "")
        {
            string documentId = "protectedDocumentId";
            if (actioName == "protectDocumentFields")
            {

                RichEditDocumentServer documentServer = new RichEditDocumentServer();
                documentServer.CalculateDocumentVariable += (s, e) =>
                {
                    if (e.VariableName == "SomeField1")
                    {
                        e.Value = "CALCULATED FIELD 1";
                        e.Handled = true;
                    }
                    if (e.VariableName == "SomeField2")
                    {
                        e.Value = "CALCULATED FIELD 2";
                        e.Handled = true;
                    }
                };

                documentServer.LoadDocument(Server.MapPath(@"~/Documents/testDOC.doc"));

                Document document = documentServer.Document;

                ProtectDocvariableFieldsInDocument(document);

                using (MemoryStream stream = new MemoryStream())
                {
                    documentServer.SaveDocument(stream, DocumentFormat.OpenXml);
                    stream.Position = 0;

                    DocumentManager.CloseDocument(documentId);
                    return RichEditExtension.Open("RichEdit", documentId, DocumentFormat.OpenXml, () => { return stream; });
                }
            }
            else if (actioName == "updateProtectedFields")
            {
                RichEditDocumentServer documentServer = new RichEditDocumentServer();
                documentServer.CalculateDocumentVariable += (s, e) =>
                {
                    if (e.VariableName == "SomeField1")
                    {
                        e.Value = "CALCULATED FIELD 1 (UPDATED ON" + DateTime.Now.ToShortTimeString() + ")";
                        e.Handled = true;
                    }
                    if (e.VariableName == "SomeField2")
                    {
                        e.Value = "CALCULATED FIELD 2 (UPDATED ON" + DateTime.Now.ToShortTimeString() + ")";
                        e.Handled = true;
                    }
                };

                documentServer.LoadDocument(Server.MapPath(@"~/Documents/testDOC.doc"));

                Document document = documentServer.Document;

                ProtectDocvariableFieldsInDocument(document);

                using (MemoryStream stream = new MemoryStream())
                {
                    documentServer.SaveDocument(stream, DocumentFormat.OpenXml);
                    stream.Position = 0;

                    DocumentManager.CloseDocument(documentId);
                    return RichEditExtension.Open("RichEdit", documentId, DocumentFormat.OpenXml, () => { return stream; });
                }
            }

            return PartialView("_RichEditPartial");
        }

        private void ProtectDocvariableFieldsInDocument(Document document)
        {
            DocumentPosition lastNonProtectedPosition = document.Range.Start;
            bool containsProtectedRanges = false;
            RangePermissionCollection rangePermissions = document.BeginUpdateRangePermissions();
            for (int i = 0; i < document.Fields.Count; i++)
            {
                Field currentField = document.Fields[i];
                string fieldCode = document.GetText(currentField.CodeRange);
                if (fieldCode.Contains("DOCVARIABLE"))
                {
                    containsProtectedRanges = true;

                    rangePermissions.AddRange(CreateRangePermissions(currentField.Range, "Admin", "Admin"));
                    if (currentField.Range.Start.ToInt() > lastNonProtectedPosition.ToInt())
                    {
                        DocumentRange rangeAfterProtection = document.CreateRange(lastNonProtectedPosition, currentField.Range.Start.ToInt() - lastNonProtectedPosition.ToInt() - 1);
                        rangePermissions.AddRange(CreateRangePermissions(rangeAfterProtection, "User", "User"));
                    }
                    lastNonProtectedPosition = currentField.Range.End;
                }
            }

            if (document.Range.End.ToInt() > lastNonProtectedPosition.ToInt())
            {
                DocumentRange rangeAfterProtection = document.CreateRange(lastNonProtectedPosition, document.Range.End.ToInt() - lastNonProtectedPosition.ToInt() - 1);
                rangePermissions.AddRange(CreateRangePermissions(rangeAfterProtection, "User", "User"));
            }
            document.EndUpdateRangePermissions(rangePermissions);

            if (containsProtectedRanges)
            {
                document.Protect("123");
            }
        }

        private IEnumerable<RangePermission> CreateRangePermissions(DocumentRange documentRange, string groupName, string userName)
        {
            List<RangePermission> rangeList = new List<RangePermission>();
            RangePermission rp = new RangePermission(documentRange);
            rp.Group = groupName;
            rp.UserName = userName;
            rangeList.Add(rp);
            return rangeList;
        }

        public static void Document_CalculateDocumentVariable(object sender, CalculateDocumentVariableEventArgs e)
        {
            var controller = new HomeController();
            var project = controller._projects.Where(x => x.Id == selectedProject).FirstOrDefault();
            switch (e.VariableName)
            {
                case "TABLE":
                    var doc1 = new RichEditDocumentServer();
                    Table table = doc1.Document.Tables.Create(doc1.Document.Range.Start, 2, 4);

                    doc1.Document.InsertSingleLineText(table.Rows[0].Cells[0].Range.Start, "ID");
                    doc1.Document.InsertSingleLineText(table.Rows[0].Cells[1].Range.Start, "Photo");
                    doc1.Document.InsertSingleLineText(table.Rows[0].Cells[2].Range.Start, "Customer Info");
                    doc1.Document.InsertSingleLineText(table.Rows[0].Cells[3].Range.Start, "Rentals");

                    for (int i = 1; i < 2; i++)
                    {
                        doc1.Document.InsertSingleLineText(table.Rows[i].Cells[0].Range.Start, $"ID {i}");

                        string customerInfo = $"Customer Info {i}\n" +
                                              $"Address: 123 Main St, Apt {i}\n" +
                                              $"Phone: (555) 123-456{i}\n" +
                                              $"Email: customer{i}@example.com";
                        doc1.Document.InsertText(table.Rows[i].Cells[2].Range.Start, customerInfo);

                        string rentalsInfo = $"Rental {i}\n" +
                                             $"Date: 01/01/202{i}\n" +
                                             $"Amount: ${100 * i}\n" +
                                             $"Status: Active";
                        doc1.Document.InsertText(table.Rows[i].Cells[3].Range.Start, rentalsInfo);
                    }

                    for (int i = 1; i < 2; i++)
                    {
                        string imagePath = System.Web.Hosting.HostingEnvironment.MapPath($"~/Content/logo.png");
                        if (System.IO.File.Exists(imagePath))
                        {
                            using (System.IO.FileStream imageStream = new System.IO.FileStream(imagePath, System.IO.FileMode.Open))
                            {
                                var documentImageSource = DevExpress.XtraRichEdit.API.Native.DocumentImageSource.FromStream(imageStream);
                                doc1.Document.Images.Insert(table.Rows[i].Cells[1].Range.Start, documentImageSource);
                            }
                        }
                    }

                    e.Value = doc1.Document;
                    e.Handled = true;
                    doc1.Document.UpdateAllFields();
                    break;
                case "client_name":
                    
                    e.Value = project.ClientName;
                    e.Handled = true;
                    break;

                case "project_name":

                    e.Value = project.ProjectName;
                    e.Handled = true;
                    break;

                case "keystone_file_id":

                    e.Value = project.KeyStoneFileId;
                    e.Handled = true;
                    break;

                case "state_registration_id":

                    e.Value = project.StateRegistrationId;
                    e.Handled = true;
                    break;

                case "project_address":

                    e.Value = project.ProjectAddress;
                    e.Handled = true;
                    break;
            }
        }
    }
}