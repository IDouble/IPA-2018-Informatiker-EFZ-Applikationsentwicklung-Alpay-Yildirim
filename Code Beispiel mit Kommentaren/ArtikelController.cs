using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using ch.muster.se.inv.dal.Models;
using ch.muster.se.inv.bll.Interfaces;
using System.IO;
using System.Drawing;
using PagedList;
using PagedList.Mvc;
using OfficeOpenXml;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using iTextSharp.tool.xml.html;
using iTextSharp.tool.xml.parser;
using iTextSharp.tool.xml.pipeline.css;
using iTextSharp.tool.xml.pipeline.end;
using iTextSharp.tool.xml.pipeline.html;

namespace ch.muster.se.inv.web.Controllers
{
    public class ArtikelController : Controller
    {
        IArtikelService _artikelService;
        IKategorieService _kategorieService;
        IBenutzerService _benutzerService;
        IRaumService _raumService;
        IFiBuKontoService _fiBuKontoService;

        public ArtikelController(IArtikelService artikelService, IKategorieService kategorieService, IBenutzerService benutzerService, IRaumService raumService, IFiBuKontoService fiBuKontoService)
        {
            _artikelService = artikelService;
            _kategorieService = kategorieService;
            _benutzerService = benutzerService;
            _raumService = raumService;
            _fiBuKontoService = fiBuKontoService;
        }

        // Das ist die Übersicht über alle Artikel, die Parameter sind optional und werden für den Suchfilter benötigt
        // Die PagedList habe ich nach diesem Tutorial aufgebaut
        // http://www.itprotoday.com/web-development/aspnet-mvc-paging-done-perfectly
        public ActionResult Index(int? KategorieID, int? FiBuKontoID, int? RaumID, int? BenutzerID, int? ArtikelYearsStart, int? ArtikelYearsEnd, string Search, string sortOrder, int page = 1, int pagesize = 25)
        {
            // Das Searchmodel wird hier abgefüllt
            ArtikelSearchModel artikelSearchModel = new ArtikelSearchModel()
            {
                KategorieID = KategorieID,
                FiBuKontoID = FiBuKontoID,
                RaumID = RaumID,
                BenutzerID = BenutzerID,
                ArtikelYearsStart = ArtikelYearsStart,
                ArtikelYearsEnd = ArtikelYearsEnd,
                Search = Search
            };

            IEnumerable<Artikel> artikelList = _artikelService.GetArtikels(artikelSearchModel);

            ViewBag.exportLink = "/Artikel/ArtikelExport?KategorieID=" + KategorieID + "&FiBuKontoID=" + FiBuKontoID + "&RaumID=" + RaumID + "&BenutzerID=" + BenutzerID + "&ArtikelYearsStart=" + ArtikelYearsStart + "&ArtikelYearsEnd=" + ArtikelYearsEnd + "&Search=" + Search;

            ViewBag.exportLinkPDF = "/Artikel/Download?KategorieID=" + KategorieID + "&FiBuKontoID=" + FiBuKontoID + "&RaumID=" + RaumID + "&BenutzerID=" + BenutzerID + "&ArtikelYearsStart=" + ArtikelYearsStart + "&ArtikelYearsEnd=" + ArtikelYearsEnd + "&Search=" + Search;

            Session["BackLink"] = "/Artikel?KategorieID=" + KategorieID + "&FiBuKontoID=" + FiBuKontoID + "&RaumID=" + RaumID + "&BenutzerID=" + BenutzerID + "&ArtikelYearsStart=" + ArtikelYearsStart + "&ArtikelYearsEnd=" + ArtikelYearsEnd + "&Search=" + Search + "&sortOrder=" + sortOrder + "&page=" + page;

            ViewBag.KategorieID = new SelectList(_kategorieService.GetKategories().OrderBy(x => x.Bezeichnung), "ID", "Bezeichnung", KategorieID);
            ViewBag.FiBuKontoID = new SelectList(_fiBuKontoService.GetFiBuKontos().OrderBy(x => x.Nummer), "ID", "Nummer", FiBuKontoID);
            ViewBag.RaumID = new SelectList(_raumService.GetRaums().OrderBy(x => x.Bezeichnung), "ID", "Bezeichnung", RaumID);
            ViewBag.BenutzerID = new SelectList(_benutzerService.GetBenutzers().OrderBy(x => x.Fullname), "ID", "Fullname", BenutzerID);
            ViewBag.ArtikelYearsStart = ArtikelYears();
            ViewBag.ArtikelYearsEnd = ArtikelYears();

            //Sortieren der Tabelle nach dem Beispiel im angegebenen Tutorial
            //https://www.itworld.com/article/2956575/development/how-to-sort-search-and-paginate-tables-in-asp-net-mvc-5.html

            ViewBag.InventarNummerSortParam = sortOrder == "InventarNummer" ? "InventarNummer_desc" : "InventarNummer";
            ViewBag.GegenstandSortParam = sortOrder == "Gegenstand" ? "Gegenstand_desc" : "Gegenstand";
            ViewBag.KategorieSortParam = sortOrder == "Kategorie" ? "Kategorie_desc" : "Kategorie";
            ViewBag.OrtSortParam = sortOrder == "Ort" ? "Ort_desc" : "Ort";
            ViewBag.NutzerSortParam = sortOrder == "Nutzer" ? "Nutzer_desc" : "Nutzer";
            ViewBag.FiBuKontoSortParam = sortOrder == "FiBuKonto" ? "FiBuKonto_desc" : "FiBuKonto";
            ViewBag.BeschaffungsdatumSortParam = sortOrder == "Beschaffungsdatum" ? "Beschaffungsdatum_desc" : "Beschaffungsdatum";
            ViewBag.BeschaffungswertSortParam = sortOrder == "Beschaffungswert" ? "Beschaffungswert_desc" : "Beschaffungswert";

            ViewBag.CurrentSort = sortOrder;

            switch (sortOrder)
            {
                case "InventarNummer":
                    artikelList = artikelList.OrderBy(x => x.InventarNummer);
                    break;
                case "InventarNummer_desc":
                    artikelList = artikelList.OrderByDescending(x => x.InventarNummer);
                    break;
                case "Gegenstand":
                    artikelList = artikelList.OrderBy(x => x.Gegenstand);
                    break;
                case "Gegenstand_desc":
                    artikelList = artikelList.OrderByDescending(x => x.Gegenstand);
                    break;
                case "Kategorie":
                    artikelList = artikelList.OrderBy(x => x.Kategorie.Bezeichnung);
                    break;
                case "Kategorie_desc":
                    artikelList = artikelList.OrderByDescending(x => x.Kategorie.Bezeichnung);
                    break;
                case "Ort":
                    artikelList = artikelList.OrderBy(x => x.Ort);
                    break;
                case "Ort_desc":
                    artikelList = artikelList.OrderByDescending(x => x.Ort);
                    break;
                case "Nutzer":
                    artikelList = artikelList.OrderBy(x => x.Nutzer.Fullname);
                    break;
                case "Nutzer_desc":
                    artikelList = artikelList.OrderByDescending(x => x.Nutzer.Fullname);
                    break;
                case "FiBuKonto":
                    artikelList = artikelList.OrderBy(x => x.Kategorie.FiBuKonto.Nummer);
                    break;
                case "FiBuKonto_desc":
                    artikelList = artikelList.OrderByDescending(x => x.Kategorie.FiBuKonto.Nummer);
                    break;
                case "Beschaffungsdatum":
                    artikelList = artikelList.OrderBy(x => x.Beschaffungsdatum);
                    break;
                case "Beschaffungsdatum_desc":
                    artikelList = artikelList.OrderByDescending(x => x.Beschaffungsdatum);
                    break;
                case "Beschaffungswert":
                    artikelList = artikelList.OrderBy(x => x.Beschaffungswert);
                    break;
                case "Beschaffungswert_desc":
                    artikelList = artikelList.OrderByDescending(x => x.Beschaffungswert);
                    break;
                default:
                    break;
            }

            ViewBag.ExportButtons = artikelList.Count() == 0 ? "disabled" : "";  // Wenn es keine Einträge in der artikelList hat, sind die Buttons deaktiviert

            return View(new PagedList<Artikel>(artikelList.ToList(), page, pagesize));
        }

        // Diese Methode wird benötigt zum erstellen der Excel Liste
        public ActionResult ArtikelExport(int? KategorieID, int? FiBuKontoID, int? RaumID, int? BenutzerID, int? ArtikelYearsStart, int? ArtikelYearsEnd, string Search)
        {
            // Das Searchmodel wird hier abgefüllt
            ArtikelSearchModel artikelSearchModel = new ArtikelSearchModel()
            {
                KategorieID = KategorieID,
                FiBuKontoID = FiBuKontoID,
                RaumID = RaumID,
                BenutzerID = BenutzerID,
                ArtikelYearsStart = ArtikelYearsStart,
                ArtikelYearsEnd = ArtikelYearsEnd,
                Search = Search
            };

            IEnumerable<Artikel> artikelList = _artikelService.GetArtikels(artikelSearchModel);

            MemoryStream ms = ArtikelToExcel(artikelList.OrderBy(d => d.InventarNummer).ToList());
            ms.WriteTo(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.ContentEncoding = new System.Text.UTF8Encoding(); // UTF-8
            Response.AddHeader("Content-Disposition", "attachment;filename=InventarListe_" + GetTimestamp(DateTime.Now) + ".xlsx");
            Response.StatusCode = 200;
            Response.End();

            return null;
        }

        // Diese Methode liefert das Excel File
        internal MemoryStream ArtikelToExcel(List<Artikel> ArtikelList)
        {
            MemoryStream Result = new MemoryStream();
            ExcelPackage pack = new ExcelPackage();
            ExcelWorksheet ws = pack.Workbook.Worksheets.Add("Inventar");

            int row = 1;

            ws.Cells[row, 1].Value = "Inventar Nummer";
            ws.Cells[row, 2].Value = "Bezeichnung";
            ws.Cells[row, 3].Value = "Beschreibung";
            ws.Cells[row, 4].Value = "Kategorie";
            ws.Cells[row, 5].Value = "Stockwerk";
            ws.Cells[row, 6].Value = "Raum";
            ws.Cells[row, 7].Value = "Nutzer";
            ws.Cells[row, 8].Value = "FiBo Konto (Nummer)";
            ws.Cells[row, 9].Value = "Beschaffungsdatum";
            ws.Cells[row, 10].Value = "Beschaffungswert (in CHF)";


            ws.Row(row).Style.Font.Size = 9;
            ws.Row(row).Style.Font.Bold = true;
            ws.Row(row).Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            ws.Row(row).Height = 18;
            ws.Row(row).Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            ws.Row(row).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            ws.Row(row).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(51, 101, 155));

            row++;

            foreach (Artikel artikel in ArtikelList)
            {
                ws.Cells[row, 1].Value = artikel.InventarNummer;
                ws.Cells[row, 2].Value = artikel.Bezeichnung;
                ws.Cells[row, 3].Value = artikel.Beschreibung;
                ws.Cells[row, 4].Value = artikel.Kategorie.Bezeichnung;
                ws.Cells[row, 5].Value = artikel.Raum.Lokalisierung;
                ws.Cells[row, 6].Value = artikel.Raum.Bezeichnung;
                ws.Cells[row, 7].Value = artikel.Nutzer.Fullname;
                ws.Cells[row, 8].Value = artikel.Kategorie.FiBuKonto.Nummer.ToString() + " ";
                ws.Cells[row, 9].Value = artikel.Beschaffungsdatum;
                ws.Cells[row, 9].Style.Numberformat.Format = @"d/m/yyyy";
                ws.Cells[row, 10].Value = artikel.Beschaffungswert;
                ws.Cells[row, 10].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";

                ws.Row(row).Style.Font.Size = 9;
                ws.Row(row).Height = 18;
                ws.Row(row).Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                row++;
            }

            ws.Cells.AutoFitColumns(1, 80);

            pack.SaveAs(Result);
            return Result;
        }

        // Methode für den PDF Download eines Artikels
        public ActionResult Download(int? KategorieID, int? FiBuKontoID, int? RaumID, int? BenutzerID, int? ArtikelYearsStart, int? ArtikelYearsEnd, string Search)
        {
            // Das Searchmodel wird hier abgefüllt
            ArtikelSearchModel artikelSearchModel = new ArtikelSearchModel()
            {
                KategorieID = KategorieID,
                FiBuKontoID = FiBuKontoID,
                RaumID = RaumID,
                BenutzerID = BenutzerID,
                ArtikelYearsStart = ArtikelYearsStart,
                ArtikelYearsEnd = ArtikelYearsEnd,
                Search = Search
            };

            IEnumerable<Artikel> artikelList = _artikelService.GetArtikels(artikelSearchModel);

            Response.ContentEncoding = new System.Text.UTF8Encoding(); // UTF-8

            Response.ContentType = "application/pdf";

            FileContentResult fcr = new FileContentResult(RenderViewToPDF("Artikel", "exportpdf", artikelList.ToArray()), "application/pdf");
            fcr.FileDownloadName = "LabelsInventar_" + GetTimestamp(DateTime.Now) + ".pdf";
            return fcr;
        }

        private byte[] RenderViewToPDF(string ControllerName, string ViewName, IEnumerable<Artikel> artikelList)
        {
            var test = ViewName;

            return RenderHtmlToPDF(RenderViewToHTML(ControllerName, ViewName, artikelList));
        }


        protected string RenderViewToHTML(string ControllerName, string ViewName, object Model)
        {
            string Html = string.Empty;
            ViewData.Model = Model;

            if (!string.IsNullOrEmpty(ControllerName))
            {
                ControllerContext.RouteData.Values["controller"] = ControllerName;
            }

            using (var sw = new StringWriter())
            {
                ViewEngineResult vr = ViewEngines.Engines.FindPartialView(ControllerContext, ViewName);
                vr.View.Render(new ViewContext(ControllerContext, vr.View, ViewData, TempData, sw), sw);
                vr.ViewEngine.ReleaseView(ControllerContext, vr.View);
                Html += sw.GetStringBuilder().ToString();
            }
            return Html;
        }

        protected byte[] RenderHtmlToPDF(string Html)
        {
            // http://stackoverflow.com/questions/36180131/using-itextsharp-xmlworker-to-convert-html-to-pdf-and-write-text-vertically
            // http://stackoverflow.com/questions/20488045/change-default-font-and-fontsize-in-pdf-using-itextsharp

            Document document = new Document(PageSize.A4, 50f, 30f, 40f, 90f);

            if (Html.Contains("class=\"landscape\""))
            {
                document.SetPageSize(iTextSharp.text.PageSize.A4.Rotate());
            }

            MemoryStream stream = new MemoryStream();
            TextReader reader = new StringReader(Html);
            PdfWriter writer = PdfWriter.GetInstance(document, stream);
            document.AddTitle("muster ag");

            XMLWorkerFontProvider fonts = new XMLWorkerFontProvider();
            CssAppliers appliers = new CssAppliersImpl(fonts);

            HtmlPipelineContext context = new HtmlPipelineContext(appliers);
            context.SetAcceptUnknown(true);
            context.SetTagFactory(Tags.GetHtmlTagProcessorFactory());

            PdfWriterPipeline pdfpipeline = new PdfWriterPipeline(document, writer);
            HtmlPipeline htmlpipeline = new HtmlPipeline(context, pdfpipeline);

            var resolver = XMLWorkerHelper.GetInstance().GetDefaultCssResolver(false);
            resolver.AddCssFile(Server.MapPath("~/Content/inv.pdf.css"), true);
            CssResolverPipeline csspipeline = new CssResolverPipeline(resolver, htmlpipeline);

            XMLWorker worker = new XMLWorker(csspipeline, true);
            XMLParser parser = new XMLParser(worker);

            document.Open();
            parser.Parse(reader);
            worker.Close();
            document.Close();

            return stream.ToArray();
        }

        // Diese Methode gibt einen Zeitstempfel zurück
        public static String GetTimestamp(DateTime value)
        {
            return value.ToString("yyyy_MM_dd_HH_mm");
        }

        public SelectList ArtikelYears()
        {

            IEnumerable<int> artikelListDates = _artikelService.GetArtikels().OrderBy(x => x.Beschaffungsdatum).Select(d => d.Beschaffungsdatum.Year).Distinct().ToList();

            List<String> dateListYear = new List<String>();

            foreach (int artikelDate in artikelListDates)
            {
                dateListYear.Add(artikelDate.ToString());
            }

            var min = artikelListDates.Min();
            var max = artikelListDates.Max();
            var difference = max - min;

            for (int i = 1; i < difference; i++)
            {
                dateListYear.Add((min + i).ToString());
            }

            SelectList artikelYears = new SelectList(dateListYear.Distinct().OrderBy(d => d).ToList());

            return artikelYears;
        }

        // Hier sieht man die Details zu einem jeweiligen Artikel
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Artikel artikel = _artikelService.GetArtikel(id);
            if (artikel == null)
            {
                return HttpNotFound();
            }
            return View(artikel);
        }

        // Diese Methode ist zum Erstellen eines Artikels nötig
        public ActionResult Create()
        {
            ViewBag.KategorieID = new SelectList(_kategorieService.GetKategories().OrderBy(x => x.Bezeichnung), "ID", "Bezeichnung");
            ViewBag.BenutzerIDNutzer = new SelectList(_benutzerService.GetBenutzers().OrderBy(x => x.Fullname), "ID", "Fullname");
            ViewBag.RaumID = new SelectList(_raumService.GetRaums().OrderBy(x => x.Bezeichnung), "ID", "Bezeichnung");
            return View();
        }

        // Diese Methode ist für den Post zuständig, die zum Erstellen eines Artikels gebraucht wird
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "ID,KategorieID,RaumID,BenutzerIDNutzer,InventarNummer,Bezeichnung,Beschreibung,Beschaffungswert,Beschaffungsdatum")] Artikel artikel)
        {
            if (ModelState.IsValid)
            {
                _artikelService.CreateArtikel(artikel);
                return RedirectToAction("Index");
            }

            ViewBag.KategorieID = new SelectList(_kategorieService.GetKategories().OrderBy(x => x.Bezeichnung), "ID", "Bezeichnung", artikel.KategorieID);
            ViewBag.BenutzerIDNutzer = new SelectList(_benutzerService.GetBenutzers().OrderBy(x => x.Fullname), "ID", "Fullname", artikel.BenutzerIDNutzer);
            ViewBag.RaumID = new SelectList(_raumService.GetRaums().OrderBy(x => x.Bezeichnung), "ID", "Bezeichnung", artikel.RaumID);
            return View(artikel);
        }

        // Diese Methode wird zum Editieren eines Artikels benötigt
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Artikel artikel = _artikelService.GetArtikel(id);
            if (artikel == null)
            {
                return HttpNotFound();
            }
            ViewBag.KategorieID = new SelectList(_kategorieService.GetKategories().OrderBy(x => x.Bezeichnung), "ID", "Bezeichnung", artikel.KategorieID);
            ViewBag.BenutzerIDNutzer = new SelectList(_benutzerService.GetBenutzers().OrderBy(x => x.Fullname), "ID", "Fullname", artikel.BenutzerIDNutzer);
            ViewBag.RaumID = new SelectList(_raumService.GetRaums().OrderBy(x => x.Bezeichnung), "ID", "Bezeichnung", artikel.RaumID);
            return View(artikel);
        }

        // Diese Methode ist für den Post zuständig, die zum Editieren eines Artikels gebraucht wird
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "ID,KategorieID,RaumID,BenutzerIDNutzer,InventarNummer,Bezeichnung,Beschreibung,Beschaffungswert,Beschaffungsdatum")] Artikel artikel)
        {
            if (ModelState.IsValid)
            {
                _artikelService.EditArtikel(artikel);
                return RedirectToAction("Index");
            }
            ViewBag.KategorieID = new SelectList(_kategorieService.GetKategories().OrderBy(x => x.Bezeichnung), "ID", "Bezeichnung", artikel.KategorieID);
            ViewBag.BenutzerIDNutzer = new SelectList(_benutzerService.GetBenutzers().OrderBy(x => x.Fullname), "ID", "Fullname", artikel.BenutzerIDNutzer);
            ViewBag.RaumID = new SelectList(_raumService.GetRaums().OrderBy(x => x.Bezeichnung), "ID", "Bezeichnung", artikel.RaumID);
            return View(artikel);
        }

        // Diese Methode wird zum Löschen eines Artikels benötigt
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Artikel artikel = _artikelService.GetArtikel(id);
            if (artikel == null)
            {
                return HttpNotFound();
            }
            return View(artikel);
        }

        // Diese Methode ist für den Post zuständig, der zum Löschen eines Artikels gebraucht wird
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            _artikelService.DeleteArtikel(id);
            return RedirectToAction("Index");
        }
    }
}