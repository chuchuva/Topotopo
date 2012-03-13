<%--
  Import from Word document into ScrewTurn Wiki
  Version 3
  http://chuchuva.com/software/screwturn-wiki-import-from-word/
  License is open source: GNU and MIT.
--%>

<%@ Page Title="Import from Word" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true"
    CodeFile="ImportWord.aspx.cs" Inherits="ScrewTurn.Wiki.ImportWord" %>

<asp:Content ID="Content1" ContentPlaceHolderID="CphMaster" runat="server">
    <h1 class="pagetitlesystem">
        <asp:Literal ID="lblImport" runat="server" Text="Import from Word" /></h1>
    <p>
        <asp:Literal ID="lblImportDescription" runat="server" Text="Import page from Word document." /></p>
    <br />
    <br />
    Enter page title:<br />
    <asp:TextBox ID="sPageName" runat="server" Width="20em" />
    <asp:RequiredFieldValidator ID="RequiredFieldValidator1" Text="Enter page name" runat="server"
        ControlToValidate="sPageName" ForeColor="Red" Display="Dynamic" />
    <asp:Label ID="lblPageNotOverwritable" Text="The page was not marked as overwritable. Update the page content to include &amp;lt;!--Overwritable--&amp;gt; in the body."
        runat="server" ForeColor="Red" Visible="false" />
    <asp:Label ID="lblAccessDenied" Text="Access was denied to the page. You may not have the edit page permission." runat="server" ForeColor="Red" Visible="false" />
    <br />
    <br />
    Select Word document:<br />
    <asp:FileUpload ID="fileUpload" runat="server" />
    <asp:RequiredFieldValidator ID="RequiredFieldValidator2" Text="Select document" runat="server"
        ControlToValidate="fileUpload" ForeColor="Red" Display="Dynamic" />
    <br />
    <br />
    <asp:Button ID="btnImport" Text="Import from this document" runat="server" OnClick="btnImport_Click" />
    <p style="margin: 3em 0 2em 0">
        Word document should have docx extension (Office 2007 and higher).</p>
    <asp:Label ID="litError" runat="server" ForeColor="Red" Visible="false" />
</asp:Content>
