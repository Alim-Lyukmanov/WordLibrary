using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

public class WordTextPerformer
{
    private string wordPath;
    private Application app;
    private Document doc;
    private Microsoft.Office.Interop.Word.Range rng;

    public string WordPath { get { return wordPath; } set { wordPath = value; } }

    public WordTextPerformer(string wordPath)
    {
        WordPath = wordPath; ;
        InitApplication();
    }

    private void InitApplication()
    {
        app = new Application();
        doc = app.Documents.Open(this.wordPath);
        app.Visible = false;
        rng = GetDocRange();
    }

    private Microsoft.Office.Interop.Word.Range GetDocRange()
    {
        var start = doc.Content.Start;
        var end = doc.Content.End;
        Microsoft.Office.Interop.Word.Range rng = doc.Range(start, end);
        return rng;
    }

    public void SetTextSize(string fontName, int size, string alligment)
    {
        rng.Font.Size = size;
    }

    public void SetTextFont(string fontName)
    {
        rng.Font.Name = fontName;
    }

    public void SetTextAlligment(string alligment)
    {
        SelectAlligment(alligment);
    }

    public void SetTextFontSizeAllignment(string fontName = "", int size = 12, string alligment = "")
    {
        rng.Font.Size = size;
        if (fontName != string.Empty)
            rng.Font.Name = fontName;
        if (alligment != string.Empty)
            SelectAlligment(alligment);
    }

    private void SelectAlligment(string alligment)
    {
        switch (alligment.ToLower())
        {
            default:
                break;
            case "center":
                rng.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                break;
            case "left":
                rng.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                break;
            case "right":
                rng.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                break;
            case "justify":
                rng.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                break;
        }
    }

    public void RemoveComments()
    {
        doc.RemoveDocumentInformation(WdRemoveDocInfoType.wdRDIComments);
    }

    public void RemoveHeaders()
    {
        rng.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Delete();
    }

    public void RemoveFooters()
    {
        rng.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Delete();
    }

    public void RemoveHeadersAndFooters()
    {
        RemoveHeaders();
        RemoveFooters();
    }

    public void DeleteSpecificPage(int pageIndex)
    {
        Microsoft.Office.Interop.Word.Range rngPage = doc.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToAbsolute, pageIndex, Type.Missing);
        rngPage.Bookmarks[@"\Page"].Range.Delete();
    }

    public void DeleteSpecificPages(List<int> pageIndexes)
    {
        foreach (int pIndex in pageIndexes)
            DeleteSpecificPage(pIndex);
    }

    public void CloseApp()
    {
        app.Quit();
    }

    public void CloseDoc()
    {
        doc.Save();
        doc.Close();
    }

    public void HighlightText(string textToFind)
    {
        InitCommonSearchSettings();
        app.Selection.Find.Text = textToFind;
        var defaultHighlightColorIndex = app.Options.DefaultHighlightColorIndex;
        SettingsToHighlight();
        app.Selection.Find.Execute(Replace: Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll);
        app.Options.DefaultHighlightColorIndex = defaultHighlightColorIndex;
    }

    private void InitCommonSearchSettings()
    {
        app.Selection.Find.ClearFormatting();
        app.Selection.Find.Replacement.ClearFormatting();
        app.Selection.Find.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;
    }

    private void SettingsToHighlight()
    {
        app.Options.DefaultHighlightColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdYellow;
        app.Selection.Find.Replacement.Highlight = 1;
    }

    private void SettingsToReplace()
    {
        app.Selection.Find.Replacement.Highlight = 0;
    }

    public void ReplaceSpecificText(string textToReplace, string replacement)
    {
        InitCommonSearchSettings();
        SettingsToReplace();
        app.Selection.Find.Text = textToReplace;
        app.Selection.Find.Replacement.Text = replacement;
        app.Selection.Find.Execute(Replace: Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll);
    }

    public void RemoveFormattedText()
    {
        Microsoft.Office.Interop.Word.Range wRange;
        for (int i = 1; i < doc.Characters.Count; i++)
        {
            var start = doc.Characters[i].Start;
            var end = doc.Characters[i].End;
            wRange = doc.Range(start, end);
            var wColor = wRange.Font.Color;
            if (wColor == WdColor.wdColorRed)
            {
                i--;
                RemoveSingleCharacter(wRange);
            }
        }
    }

    private void RemoveSingleCharacter(Microsoft.Office.Interop.Word.Range wRange)
    {
        wRange.Find.ClearFormatting();
        wRange.Find.Replacement.ClearFormatting();
        wRange.Find.Font.Color = WdColor.wdColorRed;
        wRange.Find.Format = true;
        wRange.Find.Text = wRange.Text.Trim();
        wRange.Find.Replacement.Text = string.Empty;
        wRange.Find.Forward = true;
        wRange.Find.Wrap = WdFindWrap.wdFindStop;
        wRange.Find.Execute(Replace: WdReplace.wdReplaceAll);
    }
}

