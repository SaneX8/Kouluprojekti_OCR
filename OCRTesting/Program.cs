using System;
using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using Tesseract;
using IronOcr;
//using System.Data;
//using static System.Net.Mime.MediaTypeNames;
using excel = Microsoft.Office.Interop.Excel;
//using System.IO;
using System.Threading;
//using System.Text.RegularExpressions;
//using Microsoft.Office.Interop.Excel;


namespace OCRTesting
{
    internal class Program
    {

        static void Main(string[] args)
        {
            //string tekstit;
            var Ocr = new IronTesseract();          //Ocr-instanssin luonti käyttäen IronOcr-pakettia
            Ocr.Language = OcrLanguage.FinnishBest;
            //Ocr.AddSecondaryLanguage(OcrLanguage.Financial);
            //Ocr.Configuration.PageSegmentationMode = TesseractPageSegmentationMode.SingleChar;
            //Kielimääritys Ocr:lle
            //Ocr.Configuration.TesseractVersion = TesseractVersion.Tesseract5;

            excel.Application oXL;       //Excel työkalujen määrittämistä
            excel.Workbook oWB;
            excel.Worksheet oSheet;
            excel.Range oRange;
            object misvalue = System.Reflection.Missing.Value;




            List<Row> rows = new List<Row>();  //Luodaan lista jonne tallennamme Ocr-lukijan tunnistetut tekstit
                                               //Jokainen luettu teksti tallennetaan omaan Row-muuttujaan


            Console.WriteLine("Siirry universaaliin ratkaisuun kirjoittamalla 'custom'. Vanha esimerkkitiedoston ratkaisun saat millä vaan muulla syötteellä");


            string check = Console.ReadLine();

            //ALLA UNIVERSAALI
            //ALLA UNIVERSAALI
            //ALLA UNIVERSAALI
            //ALLA UNIVERSAALI
            //ALLA UNIVERSAALI
            //ALLA UNIVERSAALI
            //ALLA UNIVERSAALI
            //ALLA UNIVERSAALI
            //ALLA UNIVERSAALI
            //ALLA UNIVERSAALI
            if (check == "custom")
            {

                Console.WriteLine("Anna tiedoston polku");

                string polku = Console.ReadLine();

                polku = polku.Trim(new Char[] { '\"',});


                using (var Input = new OcrInput())
                {
                    Input.AddPdf(polku);      //Pdf-tiedoston sijannin määrittely Ocr:lle
                    var Result = Ocr.Read(Input);

                    Result.SaveAsTextFile("..\\..\\Test.txt");

                    


                    rows = Methods.luoLista(Ocr, rows);
                    //luodaan lista, ja järjestetään se halutuksi

                    // Alustetaan arvoja joita käytetään edessäpäin vielä tarkempaan row-olioiden järjestämiseen ja tulostukseen
                    int rownumber = 0;
                    double firstLeft = 0;
                    int prevRow = 0;
                    int sarakeNro = 1; // Arvolla määritetään tulostettavan rivin sarake

                    int space = 0;
                    string feed = "";
                    //näitä arvoja käytetään sisennyksien tulostuksessa ensimmäisellä sarakkeella

                    Methods.printList(rows);
                    // Muuttujien tulostus konsoliin

                    Methods.printToTextFile(rows);
                    // Muuttujien tulostus tekstitiedostoon testausta varten


                    //Excel-tiedoston käsittely
                    try
                    {
                        oXL = new excel.Application();
                        oXL.Visible = true;

                        oWB = (excel.Workbook)(oXL.Workbooks.Add(""));
                        oSheet = (excel.Worksheet)oWB.ActiveSheet;

                        prevRow = 0;
                        rownumber = 0; //rownumber muuttujan avulla määritellään millä sarakkeella tekstiä tulostetaan. sitä verrataan käsiteltävän row-muttujan 'row' arvoon


                        //for-loopissa tulostamme listan row-muuttujien sisältöä Exceliin
                        //Lisäksi myös RegEx-kokeilua
                        for (int i = 1; i <= rows.Count; i++)
                        {
                            if (rownumber != rows[i].row) //HUOM: IF-LAUSEEN ALUSSA SUORITETAAN ENSIMMÄISEN SARAKKEEN TULOSTUKSET, TOISEN SARAKKEEN TULOSTUKSET ELSE-KOHDASSA
                                                          //verrataan rownumber-arvoa olion 'row'arvoon. Jos se on eri, on teksti vaihtanut riviä, joten uusi rivi aloitetaan uudelta ensimmäiseltä sarakkeelta
                            {                     

                                sarakeNro = 1; //Määritetään sarakkeeksi 1, koska tässä if lauseessa tulostetaan vain ensimmäisen sarakkeen tekijöitä

                                oSheet.Cells[sarakeNro][prevRow + 1] = rows[i].content;

                                rownumber++;

                                prevRow++;

                            }

                            else //HUOM: TÄÄLLÄ ELSE-LAUSEESSA SUORITETAAN TOISEN SARAKKEEN TULOSTUKSET
                            {

                                sarakeNro++;

                                oSheet.Cells[sarakeNro][prevRow] = rows[i].content;
                                //sisällön tulostus

                            }

                            oRange = oSheet.get_Range("A1", "H1");
                            oRange.EntireColumn.AutoFit();

                        }


                    }

                    catch (Exception e)
                    {

                    }
                }
                Console.ReadLine();
            }
            //YLLÄ UNIVERSAALI
            //YLLÄ UNIVERSAALI
            //YLLÄ UNIVERSAALI
            //YLLÄ UNIVERSAALI
            //YLLÄ UNIVERSAALI
            //YLLÄ UNIVERSAALI
            //YLLÄ UNIVERSAALI
            //YLLÄ UNIVERSAALI
            //YLLÄ UNIVERSAALI
            //YLLÄ UNIVERSAALI
            //YLLÄ UNIVERSAALI
            //YLLÄ UNIVERSAALI




            //ESIMERKKI KOODI
            //ESIMERKKI KOODI
            //ESIMERKKI KOODI
            //ESIMERKKI KOODI
            //ESIMERKKI KOODI
            //ESIMERKKI KOODI
            //ESIMERKKI KOODI
            //ESIMERKKI KOODI
            else //ESIMERKKI KOODI
            {
                using (var Input = new OcrInput())
                {
                    Input.AddPdf("..\\..\\test.pdf");      //Pdf-tiedoston sijannin määrittely Ocr:lle
                    var Result = Ocr.Read(Input);

                    Result.SaveAsTextFile("..\\..\\Test.txt");


                    rows = Methods.luoLista(Ocr, rows);
                    //luodaan lista, ja järjestetään se halutuksi

                    // Alustetaan arvoja joita käytetään edessäpäin vielä tarkempaan row-olioiden järjestämiseen ja tulostukseen
                    int rownumber = 0;
                    double firstLeft = 0;
                    int prevRow = 0;
                    int sarakeNro = 1; // Arvolla määritetään tulostettavan rivin sarake

                    int space = 0;
                    string feed = "";
                    //näitä arvoja käytetään sisennyksien tulostuksessa ensimmäisellä sarakkeella

                    Methods.printList(rows);
                    // Muuttujien tulostus konsoliin

                    Methods.printToTextFile(rows);
                    // Muuttujien tulostus tekstitiedostoon testausta varten


                    //Excel-tiedoston käsittely
                    try
                    {
                        oXL = new excel.Application();
                        oXL.Visible = true;

                        oWB = (excel.Workbook)(oXL.Workbooks.Add(""));
                        oSheet = (excel.Worksheet)oWB.ActiveSheet;

                        prevRow = 0;
                        rownumber = 0; //rownumber muuttujan avulla määritellään millä sarakkeella tekstiä tulostetaan. sitä verrataan käsiteltävän row-muttujan 'row' arvoon


                        //for-loopissa tulostamme listan row-muuttujien sisältöä Exceliin
                        //Lisäksi myös RegEx-kokeilua
                        for (int i = 1; i <= rows.Count; i++)
                        {
                            if (rownumber != rows[i].row) //HUOM: IF-LAUSEEN ALUSSA SUORITETAAN ENSIMMÄISEN SARAKKEEN TULOSTUKSET, TOISEN SARAKKEEN TULOSTUKSET ELSE-KOHDASSA
                                                          //verrataan rownumber-arvoa olion 'row'arvoon. Jos se on eri, on teksti vaihtanut riviä, joten uusi rivi aloitetaan uudelta ensimmäiseltä sarakkeelta
                            {

                                sarakeNro = 1; //Määritetään sarakkeeksi 1, koska tässä if lauseessa tulostetaan vain ensimmäisen sarakkeen tekijöitä

                                space = ExcelMethods.alustus(firstLeft, space, rows[i].left);
                                //alustetaan rivien sisennys
                                feed = ExcelMethods.välienLisäys(feed, space, rows[i].content);
                                //lisätään sisennykset alkuun

                                oSheet.Cells[sarakeNro][prevRow + 1] = feed;
                                //sisällön tulostus

                                rownumber++;
                                //Kasvatetaan rivinumeroa. Näin kun sitä verrataan seuraavassa oliossa ja olion 'row'-muuttuja on sama kuin edellinen rivi, on kyseessä sama rivi ja tulostusta jatketaan toisella sarakkeella.
                                feed = "";
                                space = 0;
                                //nollataan arvot


                                ExcelMethods.otsikonJaHintojenErottelu(oSheet, rows[i].content, prevRow, sarakeNro, rows[i].left, rows[i].page);
                                //ExcelMethods.hintojenErottelu(oSheet, rows[i].content, prevRow, sarakeNro);


                                ExcelMethods.otsikoidenHarmennus(oSheet, rows[i].content, prevRow);
                                //havaituista otsikoista taustaväri harmaaksi
                                ExcelMethods.otsikonKorostus(oSheet, rows[i].content, prevRow, sarakeNro);
                                //havaitusta otsikosta korostetaan teksti

                                prevRow++;
                            }

                            else //HUOM: TÄÄLLÄ ELSE-LAUSEESSA SUORITETAAN TOISEN SARAKKEEN TULOSTUKSET
                            {

                                sarakeNro++;

                                oSheet.Cells[sarakeNro][prevRow] = rows[i].content;
                                //sisällön tulostus

                                //virheellisten tuplahintojen erottelu
                                ExcelMethods.hintojenErottelu(oSheet, rows[i].content, prevRow, sarakeNro);
                                sarakeNro = ExcelMethods.sarakeKorotusCheck(sarakeNro, rows[i].content);

                                ExcelMethods.otsikoidenHarmennus(oSheet, rows[i].content, prevRow - 1);

                            }

                            oRange = oSheet.get_Range("A1", "H1");
                            oRange.EntireColumn.AutoFit();

                        }

                    }

                    catch (Exception e)
                    {

                    }
                }
            }

            Console.ReadLine();
        }
        //ESIMERKKI KOODI
        //ESIMERKKI KOODI
        //ESIMERKKI KOODI
        //ESIMERKKI KOODI
        //ESIMERKKI KOODI
        //ESIMERKKI KOODI
        //ESIMERKKI KOODI
        //ESIMERKKI KOODI
        //ESIMERKKI KOODI
        //ESIMERKKI KOODI
        //ESIMERKKI KOODI
        //ESIMERKKI KOODI
        //ESIMERKKI KOODI
        //ESIMERKKI KOODI

    }
}
