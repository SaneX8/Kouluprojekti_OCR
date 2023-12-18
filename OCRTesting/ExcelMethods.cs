using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace OCRTesting
{
    internal class ExcelMethods
    {
        public static int alustus(double alustusArvo, int välienMäärä, int leftArvo)
        {
            leftArvo = leftArvo - 100;
            alustusArvo = leftArvo / 30;
            välienMäärä = (int)Math.Round(alustusArvo);
            return välienMäärä;
        }

        public static string välienLisäys(string syöte, int välienMäärä, string content)
        {
            for (int j = 0; j < välienMäärä; j++)
            {
                syöte = syöte + "  ";
            };
            syöte = syöte + content;
            return syöte;
        }

        public static void otsikoidenHarmennus(Worksheet oSheet, string content, int prevRow)
        {
            Regex yTunari = new Regex("Tilinpäätös 1.10.2018 - 30.9.2019");
            Regex allCapsRg = new Regex("[^A-Z]");

            if (yTunari.IsMatch(content) == true) //RegEx-kokeilua
            {
                oSheet.Cells[1][prevRow + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                oSheet.Cells[2][prevRow + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                oSheet.Cells[3][prevRow + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);

            }

            if (allCapsRg.IsMatch(content) == false)
            {

                oSheet.Cells[1][prevRow + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                oSheet.Cells[2][prevRow + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                oSheet.Cells[3][prevRow + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                oSheet.Cells[1][prevRow + 1].Font.Bold = true;
            }

        }

        public static void otsikonKorostus(Worksheet oSheet, string content, int prevRow, int sarakeNro)
        {
            Regex tiliNimi = new Regex("^[0-9]{4} ");  //RegEx-koodi tilien otsikoiden tunnistukseen

            if (tiliNimi.IsMatch(content) == true) //suoritetaan sama kuin edellä paitsi tunnistetaan myös otsikkoa
            {
                oSheet.Cells[sarakeNro][prevRow + 1].Font.Bold = true;
            }
        }

        public static void hintojenErottelu(Worksheet oSheet, string content, int prevRow, int sarakeNro)
        {

            Regex rg = new Regex(",[0-9]{2} ");  //RegEx-koodi hintojen erottelua varten

            if (rg.IsMatch(content) == true)
            {
                string cutoff;
                cutoff = rg.Match(content).ToString();
                string[] result = rg.Split(content);
                oSheet.Cells[sarakeNro][prevRow] = result[0] + cutoff;
                oSheet.Cells[sarakeNro + 1][prevRow] = result[1];
            }
        }


        //SIVU 3 SOOPA ALKAA TÄÄLTÄ
        public static void otsikonJaHintojenErottelu(Worksheet oSheet, string content, int prevRow, int sarakeNro, int left, int page)
        {

            Regex rg = new Regex(",[0-9]{2} ");  //RegEx-koodi hintojen erottelua varten
            Regex rg2 = new Regex("[a-öA-Ö] -");


            if (rg.IsMatch(content) == true)
            {
                string cutoff;
                cutoff = rg.Match(content).ToString();
                string[] result = rg.Split(content);
                oSheet.Cells[sarakeNro][prevRow + 1] = result[0] + cutoff;
                oSheet.Cells[sarakeNro + 2][prevRow + 1] = result[1];
                if (rg.IsMatch(result[0] + cutoff) == true)
                {
                    otsikonJaHintojenErotteluPart2ElectricBoogaloo(oSheet, result[0] + cutoff, prevRow, sarakeNro, left);
                }
            }
        }
        public static void otsikonJaHintojenErotteluPart2ElectricBoogaloo(Worksheet oSheet, string content, int prevRow, int sarakeNro, int left)
        {

            Regex rg = new Regex("[a-öA-Ö] -");  //RegEx-koodi hintojen erottelua varten
            Regex rg2 = new Regex("[a-öA-Ö)] [0-9]");
            int firstLeft = 0;
            int space = 0;
            string feed = "";

            try
            {
                if (rg.IsMatch(content) == true)
                {
                    string cutoff;
                    cutoff = rg.Match(content).ToString();
                    string[] result = rg.Split(content);
                    space = alustus(firstLeft, space, left);
                    feed = välienLisäys(feed, space, result[0]);
                    oSheet.Cells[sarakeNro][prevRow + 1] = feed + cutoff.Remove(cutoff.Length - 2);
                    oSheet.Cells[sarakeNro + 1][prevRow + 1] = cutoff.Remove(0, 2) + result[1];
                    
                }
                else if (rg2.IsMatch(content) == true)
                {
                    string cutoff;
                    cutoff = rg2.Match(content).ToString();
                    string[] result = rg2.Split(content);
                    space = alustus(firstLeft, space, left);
                    feed = välienLisäys(feed, space, result[0]);
                    oSheet.Cells[sarakeNro][prevRow + 1] = feed + cutoff.Remove(cutoff.Length - 2);
                    oSheet.Cells[sarakeNro + 1][prevRow + 1] = cutoff.Remove(0, 2) + result[1];
                }
            }
            catch
            {
                oSheet.Cells[sarakeNro + 6][prevRow + 1] = "ERROR";
            }
        }

        public static int sarakeKorotusCheck(int sarakeNro, string content)
        {
            Regex rg = new Regex(",[0-9]{2} ");
            if (rg.IsMatch(content) == true)
            {
                sarakeNro++;
            }

            return sarakeNro;
        }


    }
}
