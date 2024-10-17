using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection.Emit;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Xceed.Document.NET;
using Xceed.Words.NET;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Border = Xceed.Document.NET.Border;
using BorderStyle = Xceed.Document.NET.BorderStyle;
using PdfSharp.Pdf;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using MigraDoc.DocumentObjectModel.Tables;
using static System.Collections.Specialized.BitVector32;
using MigraDoc.DocumentObjectModel.Visitors;
using System.Drawing.Text;

namespace Kwerendy
{
    public static class GenerateDoc
    {
        //tresc pisma o kwerendzie
        private static string wstepPisma(string nazwaZakladu) =>
            $"Informujemy, że w naszych zasobach archiwalnych przechowujemy dokumentację: {nazwaZakladu}.";        
        private static string trescPisma(string znakSparwy) =>
            $"W załączeniu przesyłamy wniosek z nadanym znakiem sprawy: {znakSparwy} . Aby wniosek został przyjęty, " +
            $"zgodnie z Regulaminem, należy przesłać wypełniony wniosek oraz uiścić opłatę za kwerendę " +
            $"(tj. poszukiwanie wskazanej we wniosku dokumentacji) w wysokości 90 zł w przeciągu 60 dni kalendarzowych od  {DateTime.Today.ToShortDateString()} r." +
            $" Przy dokonywaniu wpłat konieczne jest wskazywanie znaku sprawy {znakSparwy}. Orientacyjny czas realizacji kwerendy wynosi " +
            $"2,5 miesiąca, od momentu wpłynięcia opłaty na nasze konto.";
        private static string zakonczeniePisma() =>
            "Brak wpłaty w wyznaczonym terminie jest uznawany za rezygnację. " +
            "Po przeprowadzeniu kwerendy poinformujemy jakie dokumenty zostały odnalezione " +
            "oraz jaka jest opłata za usługę wydania dokumentacji.";

        //przetwarzanie danych
        private static string akapit1() => "Z uwagi na obowiązujące przepisy o ochronie danych " +
            "osobowych, zamieszczamy poniżej informację o zasadach przetwarzania Państwa " +
            "danych osobowych przekazanych na WNIOSKU  o wydanie dokumentacji oraz o " +
            "przetwarzaniu wynikającym z naszych działań związanych z kompletowaniem i " +
            "przekazaniem Państwu dokumentów z Archiwum. Uprzejmie prosimy o zapoznanie " +
            "się z informacją.";
        private static string akapitPrzetwarzanieDancyh() => "Państwa dane osobowe, podane " +
            "we WNIOSKU o wydanie zaświadczenia, w postaci: imienia, nazwiska, nazwiska " +
            "panieńskiego- jeżeli ma zastosowanie, daty urodzenia, imienia ojca i matki, " +
            "danych spółdzielni (byłego pracodawcy) i okresu zatrudnienia, danych adresowych, " +
            "numeru telefonu są  przetwarzane przez Związek Lustracyjny Spółdzielni z siedzibą " +
            "w Warszawie przy ul. Żurawiej 47. Przetwarzanie jest niezbędne do realizacji " +
            "Państwa uprawnienia do otrzymania zaświadczenia lub kopii dokumentów " +
            "kadrowo-płacowych z Archiwum. Jeżeli  kontaktują  się  Państwo z  Archiwum  " +
            "za  pomocą  poczty elektronicznej, to Państwa adres " +
            "e-mail może być wykorzystywany do komunikacji związanej z przygotowaniem " +
            "dokumentów o które Państwo wnioskowali.  Przy zapłacie przelewem, dane rachunku " +
            "bankowego będą przetwarzane w celu prowadzenia rozliczeń i archiwalnym. Pracownicy " +
            "Archiwum mogą zwrócić się do Państwa o podanie innych, dodatkowych informacji" +
            " i danych osobowych jeżeli skompletowanie dokumentacji na podstawie WNIOSKU o " +
            "wydanie zaświadczenia okaże się niemożliwe. Jeżeli upoważniają Państwo inne " +
            "osoby do odbioru dokumentacji lub innych działań, będziemy kontrolować i " +
            "przechowywać upoważnienia. Odbiorcami Państwa danych osobowych mogą być organy " +
            "publiczne (które mogą otrzymywać dane w ramach konkretnego postępowania) oraz " +
            "upoważnione przez administratora (czyli Związek Lustracyjny Spółdzielni) " +
            "osoby i podmioty świadczące usługi na jego rzecz. Jeżeli skorzystali Państwo" +
            " z możliwości upoważnienia innej osoby do odbioru dokumentacji lub prowadzeniu" +
            " korespondencji w Państwa imieniu, odbiorcami danych będą również te osoby. " +
            "Związek Lustracyjny Spółdzielni wdrożył odpowiednie środki techniczne i " +
            "organizacyjne aby przetwarzanie danych odbywało się zgodnie z RODO* i innymi " +
            "przepisami z zakresu ochrony danych osobowych. Po wydaniu dokumentacji, WNIOSEK " +
            "i Państwa dane osobowe będą przechowywane,  przez  okres wskazany dla danego " +
            "rodzaju sprawy w obowiązującym jednolitym rzeczowym wykazie akt dla archiwów" +
            " państwowych, który wynosi 20 lat od daty wydania dokumentu. Dane osobowe" +
            " przetwarzane w celu prowadzenia rozliczeń i dokumentacji księgowej będą " +
            "przechowywane przez okres wynikający z przepisów podatkowych oraz ustawy o " +
            "rachunkowości.\r\nJeżeli zajdzie konieczność rozpatrzenia reklamacji, to " +
            "niezbędne dane będą przechowywane do zakończenia tego postępowania. WNIOSEK " +
            "o wydanie zaświadczenia przechowujemy przez okres przewidziany dla przechowywania";
        private static string akapitPrawaIDalsze() => "Mają Państwo prawo żądania dostępu" +
            " do treści swoich danych oraz prawo do sprostowania nieprawidłowych danych" +
            " (nie dotyczy danych w dokumentacji archiwalnej). Mają również Państwo prawo" +
            " do usunięcia lub ograniczenia przetwarzania swoich danych oraz prawo wniesienia" +
            " sprzeciwu wobec przetwarzania, te prawa przysługują w określonych przypadkach" +
            " wskazanych w Art.17, Art.18 i Art.21 RODO. Przysługuje Państwu również prawo" +
            " do wniesienia skargi do organu nadzorczego - Prezesa Urzędu Ochrony Danych" +
            " Osobowych. Podanie danych jest niezbędne do wydania zaświadczenia lub" +
            " kopii dokumentów archiwalnych, jeżeli Państwo ich nie podadzą pracownicy" +
            " Archiwum nie będą mogli skompletować dokumentów oraz nie będą w stanie dokonać" +
            " weryfikacji czy po dokumenty zgłosiła się właściwa, uprawniona osoba, a bez " +
            "takiej pewności nie możemy wydać zaświadczenia lub kopii dokumentów.";
        private static string akapitKontakt() => "Związek Lustracyjny Spółdzielni  " +
            "ul. Żurawia 47, 00-680 Warszawa, www.zlsp.org.pl\r\nKontakt" +
            " do inspektora ochrony danych, e-mail: iod@zlsp.org.pl\r\n";


        public static void Generate(string folderPath, List<Entry> registrations)
        {
            string wordFilePath = folderPath + $"\\kw_{DateTime.Today.ToShortDateString()}";
            var doc = DocX.Create(wordFilePath);
            foreach (Entry entry in registrations)
            {
                MainPage(doc, entry);
                if (!entry.Description.Contains('@'))  //jeśli nie zawiera "@" to wygeneruj wniosek, kopert nie bo ksero się zatnie
                {
                    MainPage(doc, entry);
                    RequestPattern(doc, entry);
                }

                CreatePDF(folderPath, entry);
            }

            //stopka
            doc.DifferentOddAndEvenPages = true; // włącza różne stopki w dokumencie
            doc.AddFooters();
            var footer = doc.Footers.Odd;
            var fPara1 = footer.InsertParagraph("Regulamin dostępny na stronie www.zlsp.org.pl");
            fPara1.Font("TimesNewRoman").FontSize(8).Alignment = Alignment.both;
            var fPara2 = footer.InsertParagraph("Opłata naliczana jest na podstawie Rozporządzenia Ministra Kultury z dnia 10 lutego 2005 r. i Regulaminu ZLSP");
            fPara2.Font("TimesNewRoman").FontSize(8).Alignment = Alignment.both;

            doc.Save();
        }

        //generowanie jednej strony
        private static void MainPage(DocX doc, Entry entry)
        {
            // NAGŁÓWEK
            doc.InsertParagraph("Centralne Archiwum Spółdzielczości")
                .FontSize(15)
                .Bold()
                .Font("TimesNewRoman");

            doc.InsertParagraph("ul. Skośna 16     30-348 Kraków     tel. 12-262-01-79"
                + "\n" + "Konto: KBS 86 8591 0007 0020 0043 5189 0024")
                .FontSize(12)
                .Bold()
                .SpacingAfter(5)
                .Font("TimesNewRoman")
                .Alignment = Alignment.left;




            // ODSTĘP
            doc.InsertParagraph($"")
                .FontSize(12)
                .SpacingAfter(10)
                .Font("TimesNewRoman")
                .Alignment = Alignment.left;



            // TABELE
            var tab1 = doc.AddTable(1, 2);
            tab1.Rows[0].Cells[0].Paragraphs[0].Append("ZNAK SPRAWY: ").Font("TimesNewRoman");
            tab1.Rows[0].Cells[0].Paragraphs[0].Append(entry.Sign)
                .Bold()
                .Font("TimesNewRoman");
            tab1.Rows[0].Cells[1].Paragraphs[0].Append(DateTime.Today.ToShortDateString())
                .Font("TimesNewRoman")
                .Alignment = Alignment.right;
            ClearBorder(tab1);
            doc.InsertTable(tab1);

            doc.InsertParagraph().SpacingAfter(10);

            var tab2 = doc.AddTable(3, 2);
            tab2.Rows[1].Cells[1].Paragraphs[0]
                .Append("Sz. P." + "\n" + entry.Name + "\n" + entry.AdressL1 + "\n" + entry.AdressL2)
                .Font("TimesNewRoman");
            ClearBorder(tab2);
            doc.InsertTable(tab2);



            // ODSTĘP
            doc.InsertParagraph($"")
                .FontSize(12)
                .SpacingAfter(10)
                .Font("TimesNewRoman")
                .Alignment = Alignment.left;
            


            //TREŚĆ PISMA
            var wstep = doc.InsertParagraph(wstepPisma(entry.workplace));
                wstep.FontSize(12).Font("TimesNewRoman").Alignment = Alignment.both;
                wstep.IndentationFirstLine = 20.0f;

            var tresc = doc.InsertParagraph(trescPisma(entry.Sign));
                tresc.FontSize(12).Font("TimesNewRoman").Alignment = Alignment.both;
                tresc.IndentationFirstLine = 20.0f;

            var zakonczenie = doc.InsertParagraph(zakonczeniePisma());
                zakonczenie.FontSize(12).Font("TimesNewRoman").Alignment = Alignment.both;
                zakonczenie.IndentationFirstLine = 20.0f;

            doc.InsertParagraph().InsertPageBreakAfterSelf(); // nowa strona

            // nowa strona żeby się dwustronnie nie drukowały
            doc.InsertParagraph().InsertPageBreakAfterSelf(); 




        }

        //generowanie wniosku
        private static void RequestPattern(DocX doc, Entry entry)
        {
            int baseFontSize = 12;
            int space = 6;
            int bigSpace = 17;
            doc.InsertParagraph("WNIOSEK O WYDANIE DOKUMENTACJI")
                .FontSize(14).Font("TimesNewRoman").Bold().Alignment = Alignment.center;
            doc.InsertParagraph($"Znak sprawy: {entry.Sign}")
                .FontSize(12).Font("TimesNewRoman").Bold().Alignment = Alignment.right;

            doc.InsertParagraph("Dane personalne")
                .FontSize(baseFontSize).Font("TimesNewRoman").Bold().SpacingAfter(space);
            doc.InsertParagraph("1. Imię:")
                .FontSize(baseFontSize).Font("TimesNewRoman").SpacingAfter(space);
            doc.InsertParagraph("2. Nazwisko (obecne):")
                .FontSize(baseFontSize).Font("TimesNewRoman").SpacingAfter(space);
            doc.InsertParagraph("3. Nazwiska wcześniej używane:")
                .FontSize(baseFontSize).Font("TimesNewRoman").SpacingAfter(space);
            doc.InsertParagraph("4. Nazwisko w trakcie zatrudnienia:")
                .FontSize(baseFontSize).Font("TimesNewRoman").SpacingAfter(space);
            doc.InsertParagraph("5. Data urodzenia:")
                .FontSize(baseFontSize).Font("TimesNewRoman").SpacingAfter(space);
            doc.InsertParagraph("6. Adres zamieszkania / korespondencyjny:")
                .FontSize(baseFontSize).Font("TimesNewRoman").SpacingAfter(space);
            doc.InsertParagraph("...")
                .FontSize(baseFontSize).Font("TimesNewRoman").SpacingAfter(space);
            doc.InsertParagraph("...")
                .FontSize(baseFontSize).Font("TimesNewRoman").SpacingAfter(space);
            doc.InsertParagraph("7. Telefon kontaktowy:")
                .FontSize(baseFontSize).Font("TimesNewRoman").SpacingAfter(space);
            doc.InsertParagraph("8. Email:")
                .FontSize(baseFontSize).Font("TimesNewRoman").SpacingAfter(bigSpace);

            doc.InsertParagraph("Informacje dotyczące zatrudnienia")
                .FontSize(baseFontSize).Font("TimesNewRoman").Bold().SpacingAfter(space);
            doc.InsertParagraph("1. Pełna nazwa oraz adres zakładu:")
                .FontSize(baseFontSize).Font("TimesNewRoman").SpacingAfter(space);
            doc.InsertParagraph("...")
                .FontSize(baseFontSize).Font("TimesNewRoman").SpacingAfter(space);
            doc.InsertParagraph("...")
                .FontSize(baseFontSize).Font("TimesNewRoman").SpacingAfter(space);
            doc.InsertParagraph("2. Okres zatrudnienia:")
                .FontSize(baseFontSize).Font("TimesNewRoman").SpacingAfter(space);
            doc.InsertParagraph("3. Stanowisko zajmowane w okresie zatrudnienia")
                .FontSize(baseFontSize).Font("TimesNewRoman").SpacingAfter(bigSpace);

            doc.InsertParagraph("Rodzaj poszukiwanej dokumentacji")
                .FontSize(baseFontSize).Font("TimesNewRoman").Bold().SpacingAfter(space);
            doc.InsertParagraph("[ ] Świadectwo pracy")
                .FontSize(baseFontSize).Font("TimesNewRoman").SpacingAfter(space);
            doc.InsertParagraph("[ ] dokumentacja płacowa (np. kartoteki zarobkowe, listy płac...")
                .FontSize(baseFontSize).Font("TimesNewRoman").SpacingAfter(space);
            doc.InsertParagraph("Inna (proszę podać jaka): ")
                .FontSize(baseFontSize).Font("TimesNewRoman").SpacingAfter(bigSpace);

            doc.InsertParagraph("Odbiór dokumentów")
                .FontSize(baseFontSize).Font("TimesNewRoman").Bold().SpacingAfter(space);
            doc.InsertParagraph("[ ] Odbiór osobisty")
                .FontSize(baseFontSize).Font("TimesNewRoman").SpacingAfter(space);
            doc.InsertParagraph("[ ] Wysyłka pocztowa")
                .FontSize(baseFontSize).Font("TimesNewRoman").SpacingAfter(space);
            doc.InsertParagraph("[ ] Odbiór osobisty przez osobę upoważnioną (wymagane dodatkowe upoważnienie)")
                .FontSize(baseFontSize).Font("TimesNewRoman").SpacingAfter(space);
            doc.InsertParagraph("[ ] wysyłka pocztowa do osoby upoważnionej (wymagane dodatkowe upoważnienie)")
                .FontSize(baseFontSize).Font("TimesNewRoman").SpacingAfter(bigSpace);

            doc.InsertParagraph("Oświadczam, że  zapoznałem/am się z regulaminem oraz informacją o przetwarzaniu danych osobowych. " +
                "Oświadczam, że podane przeze mnie dane są prawdziwe i jestem osobą uprawnioną do otrzymania dokumentacji.")
                .FontSize(12).Font("TimesNewRoman").SpacingAfter(30).Alignment = Alignment.both;

            var tab1 = doc.AddTable(2, 4);
            tab1.Rows[0].Cells[0].Paragraphs[0].Append("....................,").Font("TimesNewRoman").Alignment = Alignment.center;
            tab1.Rows[0].Cells[1].Paragraphs[0].Append("dnia ....................").Font("TimesNewRoman").Alignment = Alignment.center;
            tab1.Rows[0].Cells[3].Paragraphs[0].Append("....................").Font("TimesNewRoman").Alignment = Alignment.center;

            tab1.Rows[1].Cells[0].Paragraphs[0].Append("Miejscowość").FontSize(10).Font("TimesNewRoman").Alignment = Alignment.center;
            tab1.Rows[1].Cells[1].Paragraphs[0].Append("Data").FontSize(10).Font("TimesNewRoman").Alignment = Alignment.center;
            tab1.Rows[1].Cells[3].Paragraphs[0].Append("Odręczny podpis").FontSize(10).Font("TimesNewRoman").Alignment = Alignment.center;
            ClearBorder(tab1);
            doc.InsertTable(tab1);



            // Przetwazranie danych osobowych
            doc.InsertParagraph("Informacja o przetwarzaniu danych osobowych")
                .FontSize(baseFontSize).Font("TimesNewRoman").UnderlineStyle(UnderlineStyle.singleLine);
            doc.InsertParagraph( akapit1() )
                .FontSize(11).Font("TimesNewRoman");

            doc.InsertParagraph("Przetwarzanie danych")
                .FontSize(baseFontSize).Font("TimesNewRoman").UnderlineStyle(UnderlineStyle.singleLine);
            doc.InsertParagraph(akapitPrzetwarzanieDancyh())
                .FontSize(11).Font("TimesNewRoman");

            doc.InsertParagraph("Prawa i dalsze informacje")
                .FontSize(baseFontSize).Font("TimesNewRoman").UnderlineStyle(UnderlineStyle.singleLine);
            doc.InsertParagraph(akapitPrawaIDalsze())
                .FontSize(11).Font("TimesNewRoman");

            doc.InsertParagraph("Dane kontaktowe")
                .FontSize(baseFontSize).Font("TimesNewRoman").UnderlineStyle(UnderlineStyle.singleLine);
            doc.InsertParagraph(akapitKontakt())
                .FontSize(11).Font("TimesNewRoman");



            doc.InsertParagraph().InsertPageBreakAfterSelf(); // nowa strona
        }

        // Generowanie pdf
        private static void CreatePDF(string folderPath, Entry entry)
        {
            string pdfFilePath = folderPath + $"\\kw_{entry.Sign}.pdf";
            MigraDoc.DocumentObjectModel.Document doc = new MigraDoc.DocumentObjectModel.Document();
            MigraDoc.DocumentObjectModel.Section section = doc.AddSection();

            Style style = doc.Styles["Normal"];
            style.Font.Name = "Times New Roman";
            style.Font.Size = 12;

            MigraDoc.DocumentObjectModel.Paragraph par1 = section.AddParagraph();



            // NAGŁÓWEK
            par1.AddText("Centralne Archiwum Społdzielczości");
            par1.Format.Font.Size = 15;
            par1.Format.Font.Bold = true;

            MigraDoc.DocumentObjectModel.Paragraph par2 = section.AddParagraph();
            par2.AddText("ul. Skośna 16     30-348 Kraków     tel. 12-262-01-79"
                + "\n" + "Konto: KBS KBS 86 8591 0007 0020 0043 5189 0024");



            MigraDoc.DocumentObjectModel.Paragraph parBlank0 = section.AddParagraph();
            parBlank0.Format.SpaceAfter = 10;


            // TABELA - ZNAK I DATA
            MigraDoc.DocumentObjectModel.Tables.Table tab = section.AddTable();
            tab.Borders.Width = 0;

            Column column00 = tab.AddColumn(Unit.FromCentimeter(8));
            Column column01 = tab.AddColumn(Unit.FromCentimeter(8));

            MigraDoc.DocumentObjectModel.Tables.Row row00 = tab.AddRow();
            MigraDoc.DocumentObjectModel.Paragraph tabPar1 = row00.Cells[0].AddParagraph($"ZNAK SPRAWY: {entry.Sign}");
            MigraDoc.DocumentObjectModel.Paragraph tabPar2 = row00.Cells[1].AddParagraph($"{DateTime.Today.ToShortDateString()}");
            tabPar2.Format.Alignment = ParagraphAlignment.Right;



            // ODSTĘP
            MigraDoc.DocumentObjectModel.Paragraph parBlank1 = section.AddParagraph();
            parBlank1.Format.SpaceAfter = 10;



            // TABELA - ADRESAT
            MigraDoc.DocumentObjectModel.Tables.Table table = section.AddTable();
            table.Borders.Width = 0;

            Column column10 = table.AddColumn(Unit.FromCentimeter(8));
            Column column11 = table.AddColumn(Unit.FromCentimeter(8));

            MigraDoc.DocumentObjectModel.Tables.Row row0 = table.AddRow();
            row0.Cells[1].AddParagraph("Sz. P.");
            row0.Cells[1].AddParagraph($"{entry.Name}");

            MigraDoc.DocumentObjectModel.Tables.Row row1 = table.AddRow();
            MigraDoc.DocumentObjectModel.Tables.Row row2 = table.AddRow();


            // ODSTĘP
            MigraDoc.DocumentObjectModel.Paragraph parBlank2 = section.AddParagraph();
            parBlank2.Format.SpaceAfter = 10;



            // TEKST PISMA
            MigraDoc.DocumentObjectModel.Paragraph parWstep = section.AddParagraph(wstepPisma(entry.Workplace));
            parWstep.Format.FirstLineIndent = 20;
            MigraDoc.DocumentObjectModel.Paragraph parTresc = section.AddParagraph(trescPisma(entry.Sign));
            parTresc.Format.FirstLineIndent = 20;
            MigraDoc.DocumentObjectModel.Paragraph parZakonczenie = section.AddParagraph(zakonczeniePisma());
            parZakonczenie.Format.FirstLineIndent = 20;



            // ODSTĘP
            MigraDoc.DocumentObjectModel.Paragraph parBlank3 = section.AddParagraph();
            parBlank2.Format.SpaceAfter = 20;


            // PODPIS
            //MigraDoc.DocumentObjectModel.Shapes.Image image = section.AddImage("C:\\Users\\Julo\\Desktop\\IT\\Testowe pliki\\podpis.png");
            MigraDoc.DocumentObjectModel.Shapes.Image image = section.AddImage(MyTools.signatureFilePath);
            image.Width = 100;
            image.Left = 300;


            //Renderowanie
            PdfDocumentRenderer renderer = new PdfDocumentRenderer();
            renderer.Document = doc;
            renderer.RenderDocument();
            renderer.Save(pdfFilePath);

        }

        private static void ClearBorder(Xceed.Document.NET.Table tab)
        {
            var myBorder = new Border(Xceed.Document.NET.BorderStyle.Tcbs_none, 0, 0, System.Drawing.Color.White);
            foreach (var row in tab.Rows)
            {
                foreach (var cell in row.Cells)
                {
                    cell.SetBorder(TableCellBorderType.Left, myBorder);
                    cell.SetBorder(TableCellBorderType.Right, myBorder);
                    cell.SetBorder(TableCellBorderType.Top, myBorder);
                    cell.SetBorder(TableCellBorderType.Bottom, myBorder);
                }
            }
        }
    }
}
