using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace DraftKingsLineupGenerator
{
    public class NBALineUp
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

            for (var i = 0; i < playerMatrix[0].Count(); i++) //pg's
            {
                lineUp.Add(playerMatrix[0][i]);
                for (var j = 0; j < playerMatrix[1].Count(); j++) //sg's
                {
                    lineUp.Add(playerMatrix[1][j]);
                    var rb1Name = lineUp[1].Name.ToString();
                    for (var l = 0; l < playerMatrix[2].Count(); l++) //sf's 
                    {
                        lineUp.Add(playerMatrix[2][l]);
                        var rb2Name = lineUp[2].Name.ToString();
                        for (var k = 0; k < playerMatrix[3].Count(); k++) //pf's 
                        {
                            lineUp.Add(playerMatrix[3][k]);
                            var wr1Name = lineUp[3].Name.ToString();
                            for (var ll = 0; ll < playerMatrix[4].Count(); ll++) //c's
                            {
                                lineUp.Add(playerMatrix[4][ll]);
                                var wr2Name = lineUp[4].Name.ToString();
                                for (var w = 0; w < playerMatrix[5].Count(); w++) //how do I flex the guard?
                                {
                                    lineUp.Add(playerMatrix[5][w]);
                                    var wr3Name = lineUp[5].Name.ToString();
                                    for (var ii = 0; ii < playerMatrix[6].Count(); ii++) //how do I flex the forward?
                                    {
                                        lineUp.Add(playerMatrix[6][ii]);
                                        var teName = lineUp[6].Name.ToString();
                                        for (var jj = 0; jj < playerMatrix[7].Count(); jj++) //util's
                                        {
                                            lineUp.Add(playerMatrix[7][jj]);

                                                var utilName = lineUp[7].Name.ToString();
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
                                                    if (!flexTestList.Contains(utilName))
                                                    {
                                                        WriteLineupsToCSVFile();
                                                    }
                                                }
                                            lineUp.RemoveAt(7); //remove util
                                        }
                                        lineUp.RemoveAt(6); //remove forward
                                    }
                                    lineUp.RemoveAt(5); //remove guard
                                }
                                lineUp.RemoveAt(4); //remove center
                            }
                            lineUp.RemoveAt(3); //remove power forward
                        }
                        lineUp.RemoveAt(2); //remove small forward
                    }
                    lineUp.RemoveAt(1); //remove shooting guard
                }
                lineUp.RemoveAt(0); //remove point guard (emptied list)
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
