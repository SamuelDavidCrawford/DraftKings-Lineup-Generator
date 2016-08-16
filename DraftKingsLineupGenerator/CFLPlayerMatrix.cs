using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DraftKingsLineupGenerator
{
    public class CFLPlayerMatrix
    {
        private List<List<Player>> _allPlayers = new List<List<Player>> { }; //returning this
        private string _theFilePath;
        private List<Player> _qbs = new List<Player> { };
        private List<Player> _rbs = new List<Player> { };
        private List<Player> _wrs = new List<Player> { };
        private List<Player> _flexs = new List<Player> { };
        private List<Player> _dsts = new List<Player> { };
        private int _qBCutoffCost;
        private int _rBCutoffCost;
        private int _wRCutoffCost;
        private int _dSTCutoffCost;
        private Excel.Application xlApp;
        private Excel.Workbook xlWorkbook;
        private Excel._Worksheet xlWorksheet;
        private Excel.Range xlRange;
        private int rowCount;
        private Player player;

        //Methods to read the positions, Names, Salaries of each player and add to respective lists
        public List<List<Player>> BuildPlayerList(int qbCutoff, int rbCutoff, int wrCutoff, int dstCutoff, string fileNameHere)
        {
            this._qBCutoffCost = qbCutoff;
            this._rBCutoffCost = rbCutoff;
            this._wRCutoffCost = wrCutoff;
            this._dSTCutoffCost = dstCutoff;
            this._theFilePath = fileNameHere;

            //Open excel workbook (.CSV downloaded from DK website)
            OpenExcelDoc(fileNameHere);

            //Read Excel Values --> Generate Lists --> Generate Matrix
            ReadPlayerPropertiesfromExcel(rowCount);

            return _allPlayers;
        }

        private void OpenExcelDoc(string xlFilePath)
        {
            xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(_theFilePath);
            xlWorksheet = xlWorkbook.Sheets[1];
            xlRange = xlWorksheet.UsedRange;
            rowCount = xlRange.Rows.Count;
        }

        private void ReadPlayerPropertiesfromExcel(int rowCount)
        {
            //Read Excel values and set Player properties
            for (int i = 9; i <= rowCount; i++)
            {
                player = new Player();
                try
                {
                    player.Name = xlRange.Cells[i, 12].value.ToString();
                    player.Position = xlRange.Cells[i, 11].value.ToString();
                }
                catch (Exception)
                {

                    throw new ArgumentException("Can not convert Name and/or Position to strings");
                }

                try
                {
                    player.Cost = Convert.ToInt32(xlRange.Cells[i, 15].value);
                    player.ID = Convert.ToInt32(xlRange.Cells[i, 14].value);
                }
                catch (Exception)
                {

                    throw new ArgumentException("Can not convert Salary and/or ID to integer value");
                }

                GenerateCFLLists(player); //Genertate Lists
            }

            GenerateNFLMatrix(); //Generate Matrix
        }

        private void GenerateCFLLists(Player player)
        {
            //Select case to add players to lists
            switch (player.Position.ToString())
            {
                case "QB":
                    if (player.Cost >= _qBCutoffCost)
                    {
                        _qbs.Add(player);
                    }
                    break;
                case "RB":
                    if (player.Cost >= _rBCutoffCost)
                    {
                        _rbs.Add(player);
                        _flexs.Add(player);
                    }
                    break;
                case "WR":
                    if (player.Cost >= _wRCutoffCost)
                    {
                        _wrs.Add(player);
                        _flexs.Add(player);
                    }
                    break;
                case "DST":
                    if (player.Cost >= _dSTCutoffCost)
                    {
                        _dsts.Add(player);
                    }
                    break;
            }
        }

        private void GenerateNFLMatrix()
        {
            //Add lists to matrix here...
            _allPlayers.Add(_qbs);
            _allPlayers.Add(_rbs);
            _allPlayers.Add(_wrs);
            _allPlayers.Add(_flexs);
            _allPlayers.Add(_dsts);
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
