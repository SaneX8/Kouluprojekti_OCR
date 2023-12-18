using IronOcr;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OCRTesting
{
    internal class Methods
    {

        public static void printList(List<Row> rows)
        {
            foreach (Row irow in rows)
            {
                Console.WriteLine(irow.content + " ROW:" + irow.row + " PAGE:" + irow.page + " LEFT:" + irow.left + " TOP:" + irow.top);
            }
        }

        public static List<Row> luoLista(IronTesseract Ocr, List<Row> rows)
        {
            int rownumber = 0;
            int prevTop = 0;
            int prevPage = 0;

            using (var Input = new OcrInput())
            {
                Input.AddPdf("..\\..\\test.pdf");      //Pdf-tiedoston sijannin määrittely Ocr:lle
                var Result = Ocr.Read(Input);

                Result.SaveAsTextFile("..\\..\\Test.txt");

                Result.Pages.ToList().ForEach(page =>
                {
                    int pagenumer = page.PageNumber;
                    page.Blocks.ToList().ForEach(block =>
                    {
                        for (int i = 0; i < block.Lines.Length; i++)
                        {

                            rows.Add(new Row(block.Lines[i].Location.Left, block.Lines[i].Location.Top, block.Lines[i].Text, block.X, block.Y, pagenumer, 0));
                            //Jokainen luettu teksti tallennetaan omaan Row-muuttujaan, jotka tallennetaan listaan
                            //Row-muuttujalla useita arvoja kuten: sisältö, etäisyys vasemmalta ja ylhäältä, sivunumero

                        }

                    });

                });

                rows = rows.OrderBy(item => item.page).ThenBy(item => item.top).ToList();

                foreach (Row irow in rows)
                {

                    if (prevPage != irow.page)
                    {
                        prevTop = 0;
                        rownumber++;
                        prevPage = irow.page;
                    }

                    if (prevTop == 0)
                    {
                        irow.row = rownumber;
                        prevTop = irow.top;
                    }

                    if (irow.top - prevTop < 5)
                    {
                        irow.row = rownumber;
                    }
                    else
                    {
                        rownumber++;
                        irow.row = rownumber;
                        prevTop = irow.top;
                    }

                }

                rows = rows.OrderBy(item => item.page).ThenBy(item => item.row).ThenBy(item => item.left).ToList();

                return rows;
            }

        }

        public static void printToTextFile(List<Row> rows)
        {
            int prevRow = 0;
            foreach (Row irow in rows)
            {
                if (prevRow != irow.row)
                {
                    //Console.WriteLine("\n"+irow.content+" ");
                    File.AppendAllText("..\\..\\WriteText.txt", "\n" + irow.content + " ");
                    prevRow = irow.row;
                }
                else
                {
                    // Console.WriteLine(irow.content+" ");
                    File.AppendAllText("..\\..\\WriteText.txt", irow.content);
                }
            }
        }
    }
}

