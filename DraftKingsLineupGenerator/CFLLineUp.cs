using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DraftKingsLineupGenerator
{
    public class CFLLineUp
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
                for (var j = 0; j < playerMatrix[1].Count(); j++) //rb1's
                {
                    lineUp.Add(playerMatrix[1][j]);
                    var rb1Name = lineUp[1].Name.ToString();
                    for (var k = 0; k < playerMatrix[2].Count() - 1; k++) //wr1's 
                    {
                        lineUp.Add(playerMatrix[2][k]);
                        var wr1Name = lineUp[2].Name.ToString();
                        for (var ll = k + 1; ll < playerMatrix[2].Count(); ll++) //wr2's (WR2 iterator = WR1 iterator + 1)
                        {
                            lineUp.Add(playerMatrix[2][ll]);
                            var wr2Name = lineUp[3].Name.ToString();
                            for (var ii = 0; ii < playerMatrix[3].Count() - 1; ii++) //flex 1's
                            {
                                lineUp.Add(playerMatrix[3][ii]);
                                var flex1Name = lineUp[4].Name.ToString();
                                for (var jj = ii+1; jj < playerMatrix[4].Count(); jj++) //flex 2's
                                {
                                    lineUp.Add(playerMatrix[3][jj]);
                                    var flex2Name = lineUp[5].Name.ToString();
                                    for (var kk = 0; kk < playerMatrix[5].Count(); kk++) //dst 
                                    {
                                        lineUp.Add(playerMatrix[4][kk]);
                                        var flexTestList = new List<string> { rb1Name, wr1Name, wr2Name, };

                                        //Calculate Line-up cost (must be between minCost and maxCost=50000)
                                        var totalCost = 0;
                                        foreach (var player in lineUp)
                                        {
                                            totalCost = player.Cost + totalCost;
                                        }

                                        //Cost Check - if within salary range, output the lineup
                                        if (totalCost <= _maxCost && totalCost >= _minCost)
                                        {
                                            if (!flexTestList.Contains(flex1Name) && !flexTestList.Contains(flex2Name))
                                            {
                                               WriteLineupsToCSVFile();
                                            }
                                        }
                                        lineUp.RemoveAt(8); //remove defense
                                    }
                                    lineUp.RemoveAt(7); //remove flex 2
                                }
                                lineUp.RemoveAt(6); //remove flex 1
                            }
                            lineUp.RemoveAt(4); //remove wr2
                        }
                        lineUp.RemoveAt(3); //remove wr1
                    }
                    lineUp.RemoveAt(1); //remove rb1
                }
                lineUp.RemoveAt(0); //remove qb (emptied list)
            }

            xlWorkbook.SaveAs(_filePath + "-Lineups.csv");
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
