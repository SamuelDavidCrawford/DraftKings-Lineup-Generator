using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace DraftKingsLineupGenerator
{
    public class LineUp
    {
        //Set the min cost of the lineUp as you so choose
        private int _maxCost = 50000;
        private int _minCost;
        private int _lineupNumber = 2;
        private List<Player> lineUp;
        private Excel.Application xlApp;
        private Excel.Workbook xlWorkbook;
        private Excel._Worksheet xlWorksheet;
        private string _filePath;

        //Algorithm to build the possible lineups
        public void BuildLineUp(List<List<Player>> playerMatrix, int minSalary, string filePath)
        {
            this._minCost = minSalary;
            this._filePath = filePath;
            lineUp = new List<Player> { };
            //Open Excel doc
            xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(filePath);
            xlWorksheet = xlWorkbook.Sheets[1];

            for (var i = 0; i < playerMatrix[0].Count(); i++) //qb's
            {
                lineUp.Add(playerMatrix[0][i]);
                for (var j = 0; j < playerMatrix[1].Count()-1; j++) //rb1's
                {
                    lineUp.Add(playerMatrix[1][j]);
                    var rb1Name = lineUp[1].Name.ToString();
                    for (var l = j+1; l < playerMatrix[1].Count(); l++) //rb2's (RB2 iterator = RB1 iterator + 1)
                    {
                        lineUp.Add(playerMatrix[1][l]);
                        var rb2Name = lineUp[2].Name.ToString();
                        for (var k = 0; k < playerMatrix[2].Count()-2; k++) //wr1's 
                        {
                            lineUp.Add(playerMatrix[2][k]);
                            var wr1Name = lineUp[3].Name.ToString();
                            for (var ll = k+1; ll < playerMatrix[2].Count()-1; ll++) //wr2's (WR2 iterator = WR1 iterator + 1)
                            {
                                lineUp.Add(playerMatrix[2][ll]);
                                var wr2Name = lineUp[4].Name.ToString();
                                for (var w = ll+1; w < playerMatrix[2].Count(); w++) //wr3's (WR3 iterator = WR2 iterator + 1)
                                {
                                    lineUp.Add(playerMatrix[2][w]);
                                    var wr3Name = lineUp[5].Name.ToString();
                                    for (var ii = 0; ii < playerMatrix[3].Count(); ii++) //te's
                                    {
                                        lineUp.Add(playerMatrix[3][ii]);
                                        var teName = lineUp[6].Name.ToString();
                                        for (var jj = 0; jj < playerMatrix[4].Count(); jj++) //dst's
                                        {
                                            lineUp.Add(playerMatrix[4][jj]);
                                            for (var kk = 0; kk < playerMatrix[5].Count(); kk++) //flex's 
                                            {
                                                lineUp.Add(playerMatrix[5][kk]);
                                                var flexName = lineUp[8].Name.ToString();
                                                var flexTestList = new List<string> { rb1Name, rb2Name, wr1Name, wr2Name, wr3Name, teName };
                                                
                                                //Calculate Line-up cost (must be between minCost and maxCost=50000)
                                                var totalCost = 0;
                                                foreach (var player in lineUp)
                                                {
                                                    totalCost = player.Cost + totalCost;
                                                }

                                                //Cost Check - if within salary range, output the lineup
                                                if (totalCost <= _maxCost && totalCost >= _minCost)
                                                {
                                                    if (!flexTestList.Contains(flexName))
                                                    {
                                                        WriteLineupsToCSVFile();
                                                    }
                                                }
                                                lineUp.RemoveAt(8); //remove flex
                                            }
                                            lineUp.RemoveAt(7); //remove dst
                                        }
                                        lineUp.RemoveAt(6); //remove te
                                    }
                                    lineUp.RemoveAt(5); //remove wr3
                                }
                                lineUp.RemoveAt(4); //remove wr2
                            }
                            lineUp.RemoveAt(3); //remove wr1
                        }
                        lineUp.RemoveAt(2); //remove rb2
                    }
                    lineUp.RemoveAt(1); //remove rb1
                }
                lineUp.RemoveAt(0); //remove qb (emptied list)
            }

            xlWorkbook.SaveAs(@"C:\Users\Samuel Crawford\Documents\C-Sharp\DraftKings\output3.csv");
            CloseExcelDoc();

        }

        private void WriteLineupsToCSVFile()
        {
            var positionCell = 1;
            foreach (var player in lineUp)
            {
                xlWorksheet.Cells[_lineupNumber, positionCell].value = player.ID.ToString();
                positionCell = positionCell + 1;
            }
            _lineupNumber = _lineupNumber + 1;
        }

        private void CloseExcelDoc()
        {
            //Cleanup & Release
            GC.Collect();
            GC.WaitForPendingFinalizers();
            xlWorkbook.Close();
            xlApp.Quit();
        }

    }
}
