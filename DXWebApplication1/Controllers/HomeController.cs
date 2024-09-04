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
using DevExpress.Web.Internal.XmlProcessor;
using DevExpress.Web.ASPxRichEdit.Internal;
using static DevExpress.XtraPrinting.Native.ExportOptionsPropertiesNames;

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

                documentServer.LoadDocument(Server.MapPath(@"~/Documents/testDOC1.doc"));

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

                documentServer.LoadDocument(Server.MapPath(@"~/Documents/testDOC1.doc"));

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
            
            else if (actioName == "protectSection")
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

                documentServer.LoadDocument(Server.MapPath(@"~/Documents/testDOC1.doc"));

                Document document = documentServer.Document;

                ProtectSectionInDocument(document);

                using (MemoryStream stream = new MemoryStream())
                {
                    documentServer.SaveDocument(stream, DocumentFormat.OpenXml);
                    stream.Position = 0;

                    DocumentManager.CloseDocument(documentId);
                    return RichEditExtension.Open("RichEdit", documentId, DocumentFormat.OpenXml, () => { return stream; });
                }
            }
            
            else if (actioName == "insertCompTable")
            {
                RichEditDocumentServer documentServer1 = new RichEditDocumentServer();
                documentServer1.LoadDocument(Server.MapPath(@"~/Documents/testDOC1.doc"));

                var doc = documentServer1.Document;
                var position = doc.Range.End;
                doc.Fields.Create(position, " DOCVARIABLE TABLE ");

                documentServer1.CalculateDocumentVariable += (s, e) =>
                {
                    if (e.VariableName == "TABLE")
                    {
                        RichEditDocumentServer documentServer = new RichEditDocumentServer();
                        Table table = documentServer.Document.Tables.Create(documentServer.Document.Range.Start, 2, 4);

                        documentServer.Document.InsertSingleLineText(table.Rows[0].Cells[0].Range.Start, "ID");
                        documentServer.Document.InsertSingleLineText(table.Rows[0].Cells[1].Range.Start, "Photo");
                        documentServer.Document.InsertSingleLineText(table.Rows[0].Cells[2].Range.Start, "Customer Info");
                        documentServer.Document.InsertSingleLineText(table.Rows[0].Cells[3].Range.Start, "Rentals");

                        for (int i = 1; i < 2; i++)
                        {
                            documentServer.Document.InsertSingleLineText(table.Rows[i].Cells[0].Range.Start, $"ID {i}");

                            string customerInfo = $"Customer Info {i}\n" +
                                                  $"Address: 123 Main St, Apt {i}\n" +
                                                  $"Phone: (555) 123-456{i}\n" +
                                                  $"Email: customer{i}@example.com";
                            documentServer.Document.InsertText(table.Rows[i].Cells[2].Range.Start, customerInfo);

                            string rentalsInfo = $"Rental {i}\n" +
                                                 $"Date: 01/01/202{i}\n" +
                                                 $"Amount: ${100 * i}\n" +
                                                 $"Status: Active";
                            documentServer.Document.InsertText(table.Rows[i].Cells[3].Range.Start, rentalsInfo);
                        }

                        for (int i = 1; i < 2; i++)
                        {
                            string imagePath = System.Web.Hosting.HostingEnvironment.MapPath($"~/Content/logo.png");
                            if (System.IO.File.Exists(imagePath))
                            {
                                using (System.IO.FileStream imageStream = new System.IO.FileStream(imagePath, System.IO.FileMode.Open))
                                {
                                    var documentImageSource = DevExpress.XtraRichEdit.API.Native.DocumentImageSource.FromStream(imageStream);
                                    documentServer.Document.Images.Insert(table.Rows[i].Cells[1].Range.Start, documentImageSource);
                                }
                            }
                        }

                        e.Value = documentServer.Document;
                        e.Handled = true;
                    }
                };

                documentServer1.Document.UpdateAllFields();
                documentServer1.Document.UnlinkAllFields();
                Document document = doc;

                using (MemoryStream stream = new MemoryStream())
                {
                    documentServer1.SaveDocument(stream, DocumentFormat.OpenXml);
                    stream.Position = 0;

                    DocumentManager.CloseDocument(documentId);
                    return RichEditExtension.Open("RichEdit", documentId, DocumentFormat.OpenXml, () => { return stream; });
                }
            }
            
            else if(actioName == "insertComp")
            {
                RichEditDocumentServer documentServer2 = new RichEditDocumentServer();
                documentServer2.LoadDocument(Server.MapPath(@"~/Documents/testDOC1.doc"));

                var doc = documentServer2.Document;
                var position = doc.Range.Start;
                doc.Fields.Create(position, " DOCVARIABLE COMP1 ");
                doc.Fields.Create(position, " DOCVARIABLE COMP2 ");
                    
                documentServer2.CalculateDocumentVariable += (s, e) =>
                {
                    if (e.VariableName == "COMP1")
                    {
                        RichEditDocumentServer documentServer = new RichEditDocumentServer();
                        Table table = documentServer.Document.Tables.Create(documentServer.Document.Range.Start, 2, 1);

                        doc.InsertSection(doc.Range.Start);
                        doc.Sections[doc.Sections.Count - 1].StartType = SectionStartType.NextPage;

                        documentServer.Document.InsertSingleLineText(table.Rows[0].Cells[0].Range.Start, "Customer Info");

                            string customerInfo = $"Comp 1 info;\n" +
                                                  $"Address Comp2: 123 Main St, Apt \n" +
                                                  $"Phone: (555) 123-456\n" +
                                                  $"Email: customer@example.com";
                            documentServer.Document.InsertText(table.Rows[1].Cells[0].Range.Start, customerInfo);
                        
                        e.Value = documentServer.Document;
                        e.Handled = true;
                    }

                    else if(e.VariableName == "COMP2")
                    {
                        // Вставка розриву сторінки перед вставкою поля COMP2
                        doc.InsertSection(doc.Range.End);
                        doc.Sections[doc.Sections.Count - 1].StartType = SectionStartType.NextPage;

                        RichEditDocumentServer documentServer = new RichEditDocumentServer();
                        Table table = documentServer.Document.Tables.Create(documentServer.Document.Range.Start, 2, 1);

                        documentServer.Document.InsertSingleLineText(table.Rows[0].Cells[0].Range.Start, "Customer Info");

                            string customerInfo = $"Comp 2 info; 2\n" +
                                                  $"Address Comp2: 123 Main St, Apt \n" +
                                                  $"Phone: (555) 123-456\n" +
                                                  $"Email: customer2@example.com";
                            documentServer.Document.InsertText(table.Rows[1].Cells[0].Range.Start, customerInfo);

                        e.Value = documentServer.Document;
                        e.Handled = true;
                    }
                };

                documentServer2.Document.UpdateAllFields();
                //documentServer2.Document.UnlinkAllFields();
                Document document = doc;

                using (MemoryStream stream = new MemoryStream())
                {
                    documentServer2.SaveDocument(stream, DocumentFormat.OpenXml);
                    stream.Position = 0;

                    DocumentManager.CloseDocument(documentId);
                    return RichEditExtension.Open("RichEdit", documentId, DocumentFormat.OpenXml, () => { return stream; });
                }

            }
            
            return PartialView("_RichEditPartial");
        }
        public void ProtectDocvariableFieldsInDocument(Document document)
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

        public void ProtectSectionInDocument(Document document)
        {
             DocumentPosition lastNonProtectedPosition = document.Range.Start;
             RangePermissionCollection rangePermissions = document.BeginUpdateRangePermissions();
                
             Section currentSection = document.Sections[0];
             DocumentRange sectionRange = currentSection.Range;
             rangePermissions.AddRange(CreateRangePermissions( sectionRange, "User", "User"));
             

             Section currentSection1 = document.Sections[1];
             DocumentRange sectionRange1 = currentSection1.Range;
             rangePermissions.AddRange(CreateRangePermissions(sectionRange1, "Admin", "Admin"));

            document.EndUpdateRangePermissions(rangePermissions);
            document.Protect("123");
        }
    
        public IEnumerable<RangePermission> CreateRangePermissions(DocumentRange documentRange, string groupName, string userName)
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