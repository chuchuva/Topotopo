// Import from Word document into ScrewTurn Wiki
// Version 4
// http://chuchuva.com/software/screwturn-wiki-import-from-word/
// License is open source: GNU and MIT.
namespace ScrewTurn.Wiki
{
    #region using

    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Linq;
    using System.Text;
    using System.Threading;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Drawing;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
    using PluginFramework;
    using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
    using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
    using Path = System.IO.Path;
    using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
    using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
    using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
    using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
    using TableProperties = DocumentFormat.OpenXml.Wordprocessing.TableProperties;
    using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
    using TableStyle = DocumentFormat.OpenXml.Wordprocessing.TableStyle;
    using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
    using Underline = DocumentFormat.OpenXml.Wordprocessing.Underline;

    #endregion

    /// <summary>
    ///   ImportWord processes an OpenXML document and creates a WikiMarkup page.
    /// </summary>
    public partial class ImportWord : BasePage
    {
        private static int _tableindex;

        /// <summary>
        ///   Takes a Paragraph from the document and gets the styled text, then applies heading and list formatting.
        /// </summary>
        /// <param name = "p">The Paragraph to convert</param>
        /// <param name = "appendLine"></param>
        /// <returns>A String with the WikiMarkup</returns>
        private static String GetFormattedText(Paragraph p, ref Boolean appendLine)
        {
            appendLine = true;
            String text = GetStyledText(p);
            if (p.ParagraphProperties != null && p.ParagraphProperties.ParagraphStyleId != null &&
                p.ParagraphProperties.ParagraphStyleId.Val != null &&
                p.ParagraphProperties.ParagraphStyleId.Val.Value != null)
            {
                if (String.Equals(p.ParagraphProperties.ParagraphStyleId.Val.Value, "Heading1"))
                {
                    text = String.Format("== {0} ==", p.InnerText);
                }
                else if (String.Equals(p.ParagraphProperties.ParagraphStyleId.Val.Value, "Heading2"))
                {
                    text = String.Format("=== {0} ===", p.InnerText);
                }
                else if (String.Equals(p.ParagraphProperties.ParagraphStyleId.Val.Value, "Heading3"))
                {
                    text = String.Format("==== {0} ====", p.InnerText);
                }
                else if (String.Equals(p.ParagraphProperties.ParagraphStyleId.Val.Value, "Heading4"))
                {
                    text = String.Format("===== {0} =====", p.InnerText);
                }
            }

            if (p.ParagraphProperties != null && p.ParagraphProperties.NumberingProperties != null &&
                p.ParagraphProperties.NumberingProperties.NumberingId != null)
            {
                if (p.ParagraphProperties.NumberingProperties.NumberingLevelReference != null &&
                    p.ParagraphProperties.NumberingProperties.NumberingLevelReference.Val != null)
                {
                    switch (p.ParagraphProperties.NumberingProperties.NumberingLevelReference.Val.Value)
                    {
                        case 0:
                            text = String.Format("* {0}", p.InnerText);
                            break;

                        case 1:
                            text = String.Format("** {0}", p.InnerText);
                            break;

                        case 2:
                            text = String.Format("*** {0}", p.InnerText);
                            break;

                        case 3:
                            text = String.Format("**** {0}", p.InnerText);
                            break;
                    }
                    appendLine = false;
                }
            }
            return text;
        }

        private static string GetParagraphText(Paragraph p, string fullpagename, WordprocessingDocument doc,
                                               List<Blip> blips)
        {
            bool appendLine = true;
            var sb = new StringBuilder();
            sb.Append(GetFormattedText(p, ref appendLine));

            foreach (Blip blip in p.Descendants<Blip>())
            {
                OpenXmlPart imagePart = doc.MainDocumentPart.GetPartById(blip.Embed.Value);
                string fl = Path.GetFileName(imagePart.Uri.ToString());
                sb.Append(string.Format("[image||{{UP({0})}}{1}]", fullpagename, fl));
                blips.Add(blip);
            }

            if (appendLine) sb.AppendLine();
            sb.AppendLine();
            return sb.ToString();
        }

        private static string GetParagraphTextInTable(Paragraph p, string fullpagename, WordprocessingDocument doc,
                                                      List<Blip> blips)
        {
            var sb = new StringBuilder();
            sb.Append(GetStyledText(p));

            foreach (Blip blip in p.Descendants<Blip>())
            {
                OpenXmlPart imagePart = doc.MainDocumentPart.GetPartById(blip.Embed.Value);
                string fl = Path.GetFileName(imagePart.Uri.ToString());
                sb.Append(string.Format("[image||{{UP({0})}}{1}]", fullpagename, fl));
                blips.Add(blip);
            }

            return sb.ToString();
        }

        /// <summary>
        ///   Takes a Paragraph from the document and converts it to WikiMarkup
        /// </summary>
        /// <param name = "p">The Paragraph to convert</param>
        /// <returns>A String with the WikiMarkup</returns>
        private static String GetStyledText(OpenXmlElement p)
        {
            var text = new StringBuilder();
            if (p.ChildElements == null)
            {
                text.Append(p.InnerText);
            }
            else
            {
                // Create a FILO stack for the styles (First in, last out).
                // For example, if we push bold then underline, we should pop underline then bold
                // Therefore, the formatting would be: '''__text__'''
                // If we used a FIFO stack like Queue (First in, first out), we would get strange formatting
                // An example out the FIFO formatting would be: '''__text'''__
                // While this would be processed correctly in the wiki, the user would be confused by such a formatting style.
                var stack = new Stack<OpenXmlLeafElement>();
                foreach (Run ch in Enumerable.OfType<Run>(p.ChildElements))
                {
                    foreach (OpenXmlElement el in ch.ChildElements)
                    {
                        // If this is a single line break, replace it with a {BR} element (no line break afterwards for table compatability).
                        if (el is Break)
                        {
                            text.Append("{BR}");
                            continue;
                        }
                        var rp = el as RunProperties;
                        var tx = el as Text;
                        if (rp != null)
                        {
                            if (rp.Bold != null)
                            {
                                text.Append("'''");
                                stack.Push(rp.Bold);
                            }
                            if (rp.Italic != null)
                            {
                                text.Append("''");
                                stack.Push(rp.Italic);
                            }
                            if (rp.Underline != null)
                            {
                                text.Append("__");
                                stack.Push(rp.Underline);
                            }
                        }
                        if (tx != null)
                            text.Append(tx.InnerText);
                    }

                    // Clean-up the stack.
                    while (stack.Count > 0)
                    {
                        OpenXmlLeafElement ele = stack.Pop();
                        if (ele is Bold)
                            text.Append("'''");
                        if (ele is Italic)
                            text.Append("''");
                        if (ele is Underline)
                            text.Append("__");
                    }
                }
            }
            return text.ToString();
        }

        /// <summary>
        ///   Takes a Table from the document and processes each row, feeding the row's paragraph through styled text function.
        /// </summary>
        /// <param name = "t"></param>
        /// <returns></returns>
        private static String GetTableText(Table t, string fullpagename, WordprocessingDocument doc, List<Blip> blips)
        {
            // Start the table.
            TableProperties props = t.ChildElements.OfType<TableProperties>().FirstOrDefault();
            TableBorders borders = props.TableBorders;
            TableStyle style = props.TableStyle;
            string cellStyle = "";
            if (borders == null && style != null)
            {
                try
                {
                    Style myStyle =
                        doc.MainDocumentPart.StyleDefinitionsPart.Styles.OfType<Style>().Where(
                            x => x.StyleId == style.Val.Value).FirstOrDefault();
                    borders = myStyle.StyleTableProperties.TableBorders;
                }
                catch {}
            }
            if (borders != null)
            {
                if (borders.InsideHorizontalBorder != null)
                    cellStyle +=
                        string.Format(
                            "border-top-style: inset; border-top-width: {0}px; border-top-color: {1}; " +
                            "border-bottom-style: inset; border-bottom-width: {0}px; border-bottom-color: {1}; ",
                            borders.InsideHorizontalBorder.Size/4,
                            borders.InsideHorizontalBorder.Color.ToString() == "auto"
                                ? "black"
                                : "#" + borders.InsideHorizontalBorder.Color);
                if (borders.InsideVerticalBorder != null)
                    cellStyle +=
                        string.Format(
                            "border-left-style: inset; border-left-width: {0}px; border-left-color: {1}; " +
                            "border-right-style: inset; border-right-width: {0}px; border-right-color: {1}; ",
                            borders.InsideVerticalBorder.Size/4,
                            borders.InsideVerticalBorder.Color.ToString() == "auto"
                                ? "black"
                                : "#" + borders.InsideVerticalBorder.Color);
            }
            var builder = new StringBuilder();
            if (!string.IsNullOrEmpty(cellStyle))
            {
                _tableindex++;
                builder.AppendLine("<style type='text/css'>");
                builder.AppendFormat("TD.Index{0}", _tableindex);
                builder.AppendLine(" {");
                builder.AppendLine(cellStyle);
                builder.AppendLine("}");
                builder.AppendLine("</style>");
            }
            builder.Append("{|");
            builder.AppendLine();
            bool firstRow = true;
            foreach (TableRow item in Enumerable.OfType<TableRow>(t.ChildElements))
            {
                if (firstRow)
                    firstRow = false;
                else
                    builder.AppendLine("|-");
                foreach (TableCell cell in item.OfType<TableCell>())
                {
                    builder.Append("| ");
                    if (!string.IsNullOrEmpty(cellStyle))
                        builder.AppendFormat("class=\"Index{0}\" | ", _tableindex);
                    var isb = new StringBuilder();
                    foreach (Paragraph p in cell.OfType<Paragraph>())
                    {
                        string pmu = GetParagraphTextInTable(p, fullpagename, doc, blips);
                        Debug.WriteLine(pmu);
                        isb.Append(pmu);
                        isb.Append("{BR}");
                    }
                    string celldata = isb.ToString();
                    builder.AppendLine(celldata.Substring(0, celldata.LastIndexOf("{BR}")));
                }
            }
            // End the table.
            builder.AppendLine("|}");
            return builder.ToString();
        }

        protected void Page_Load(object sender, EventArgs e) {}

        protected void btnImport_Click(object sender, EventArgs e)
        {
            lblPageNotOverwritable.Visible = false;
            litError.Visible = false;
            lblAccessDenied.Visible = false;
            if (!IsValid)
                return;
            IPagesStorageProviderV30 provider;
            NamespaceInfo ns = DetectNamespaceInfo();
            provider =
                ns == null
                    ? Collectors.PagesProviderCollector.GetProvider(Settings.DefaultPagesProvider)
                    : ns.Provider;
            string pagetitle = sPageName.Text;
            string pagename = GenerateAutoName(pagetitle);
            // When referring to the page, use pagename. When referring to the title, use pagetitle.
            PageInfo page = Pages.FindPage(NameTools.GetFullName(DetectNamespace(), pagename), provider);
            string currentUser = SessionFacade.GetCurrentUsername();
            string[] currentGroups = SessionFacade.GetCurrentGroupNames();
            bool canCreateNewPages = AuthChecker.CheckActionForNamespace(ns, Actions.ForNamespaces.CreatePages,
                                                                         currentUser, currentGroups);
            bool canCreateNewPagesWithApproval = false;
            bool create = false;
            bool canEdit = false;
            bool canEditWithApproval = false;
            switch (Settings.ChangeModerationMode)
            {
                case ChangeModerationMode.RequirePageEditingPermissions:
                    canCreateNewPagesWithApproval = AuthChecker.CheckActionForNamespace(ns,
                                                                                        Actions.ForNamespaces.
                                                                                            ModifyPages, currentUser,
                                                                                        currentGroups);
                    break;
                case ChangeModerationMode.RequirePageViewingPermissions:
                    canCreateNewPagesWithApproval = AuthChecker.CheckActionForNamespace(ns,
                                                                                        Actions.ForNamespaces.ReadPages,
                                                                                        currentUser, currentGroups);
                    break;
                default:
                    canCreateNewPagesWithApproval = false;
                    break;
            }
            if (page == null)
            {
                // Page does not exist, check the permissions for creating a new page.
                if (!canCreateNewPages && !canCreateNewPagesWithApproval)
                {
                    lblAccessDenied.Visible = true;
                    return;
                }
                create = true;
            }
            else
            {
                // Page does exists, check if the page is overwritable
                PageContent pageContent = Content.GetPageContent(page, false);
                if (!pageContent.Content.Contains("<!--Overwritable-->"))
                {
                    lblPageNotOverwritable.Visible = true;
                    return;
                }
                // And then check to see if the user can write to it.
                Pages.CanEditPage(page, currentUser, currentGroups, out canEdit, out canEditWithApproval);
                if (!canEdit && !canEditWithApproval)
                {
                    lblAccessDenied.Visible = true;
                    return;
                }
            }

            try
            {
                var blips = new List<Blip>();
                WordprocessingDocument doc = WordprocessingDocument.Open(fileUpload.FileContent, false /* isEditable */);
                Body body = doc.MainDocumentPart.Document.Body;

                var sb = new StringBuilder();
                string filename = fileUpload.FileName;
                foreach (OpenXmlElement child in body.ChildElements)
                {
                    if (child is Table)
                    {
                        // Process the table.
                        var t = child as Table;
                        sb.Append(GetTableText(t, NameTools.GetFullName(DetectNamespace(), pagename), doc, blips));
                        sb.AppendLine();
                    }
                    else if (child is Paragraph)
                    {
                        // Process the paragraph.
                        var p = child as Paragraph;
                        sb.Append(GetParagraphText(p, NameTools.GetFullName(DetectNamespace(), pagename), doc, blips));
                    }
                }
                // Don't close the document yet, we still need to process the blips
                // Final checks before creating and updating the page.
                // This was moved from above the document parsing to prevent pages that failed document parsing.
                string username = (SessionFacade.LoginKey == null)
                                      ? Request.UserHostAddress
                                      : SessionFacade.CurrentUsername;
                // Default mode is backup, to save the previous revision.
                SaveMode saveMode = SaveMode.Backup;
                if (page == null)
                {
                    // If the page doesn't exist, create it.
                    Pages.CreatePage(ns, pagename, provider);
                    page = Pages.FindPage(NameTools.GetFullName(DetectNamespace(), pagename), provider);
                    // If the user can only create pages with approval,
                    //   fill the page with some generic text and then set the save mode to draft.
                    if (!canCreateNewPages)
                    {
                        Pages.ModifyPage(page, pagetitle, username, DateTime.Now, String.Empty,
                                         "{s:DraftByUser}{BR}~~~~",
                                         null, String.Empty, SaveMode.Normal);
                        saveMode = SaveMode.Draft;
                    }
                }
                else
                {
                    // If the user can only edit with approval, set the save mode to draft.
                    if (!canEdit && canEditWithApproval)
                    {
                        saveMode = SaveMode.Draft;
                    }
                }
                IFilesStorageProviderV30 filesStorage =
                    Collectors.FilesProviderCollector.GetProvider(Settings.DefaultFilesProvider);
                // Now add all the attachments.
                foreach (Blip blip in blips)
                {
                    OpenXmlPart imagePart = doc.MainDocumentPart.GetPartById(blip.Embed.Value);
                    string fl = Path.GetFileName(imagePart.Uri.ToString());
                    filesStorage.StorePageAttachment(page, fl, imagePart.GetStream(), true /* overwrite */);
                }
                doc.Close();
                // Pull the page contents back out and then write the modification.
                PageContent content = Content.GetPageContent(page, false /* cached */);
                // Note that the ModifyPage is the ABSOLUTE LAST THING that we do.
                // If we fail at something above this, nothing about the existing page has been modified.
                // Therefore we can safely exit without any harm, except until this point.
                Pages.ModifyPage(page, pagetitle, username, DateTime.Now, "Imported from " + filename,
                                 sb.ToString(), content.Keywords, content.Description, saveMode);
                UrlTools.Redirect(UrlTools.BuildUrl(Tools.UrlEncode(page.FullName), Settings.PageExtension));
            }
            catch (ThreadAbortException)
            {
                // When UrlTools.Redirect is called, the thread is aborted.
                // In order to prevent us from accidentally deleting the page below, we will catch the ThreadAbortException here.
            }
            catch (Exception ex)
            {
                litError.Text = "<code style='color:red'><pre style='color:red'>" + ex + "</code></pre>";
                litError.Visible = true;
                // If we're creating a new page and the page exists, clean up the failed attempt.
                if (page != null && create)
                {
                    try
                    {
                        Pages.DeletePage(page);
                    }
                    catch {}
                }
            }
        }

        /// <summary>
        /// Generates an automatic page name.
        /// This is Edit.GenerateAutoName from the original source - copied because it is private there.
        /// </summary>
        /// <param name="title">The page title.</param>
        /// <returns>The name.</returns>
        private static string GenerateAutoName(string title)
        {
           // Replace all non-alphanumeric characters with dashes
            if(title.Length == 0) return "";
            
            StringBuilder buffer = new StringBuilder(title.Length);
            
            foreach(char ch in title.Normalize(NormalizationForm.FormD).Replace("\"", "").Replace("'", "")) {
               var unicat = char.GetUnicodeCategory(ch);
               if(unicat == System.Globalization.UnicodeCategory.LowercaseLetter ||
                  unicat == System.Globalization.UnicodeCategory.UppercaseLetter ||
                  unicat == System.Globalization.UnicodeCategory.DecimalDigitNumber) {
                  buffer.Append(ch);
               }
               else if(unicat != System.Globalization.UnicodeCategory.NonSpacingMark) buffer.Append("-");
            }
            
            while(buffer.ToString().IndexOf("--") >= 0) {
               buffer.Replace("--", "-");
            }
            
            return buffer.ToString().Trim('-');
        }

    }
}