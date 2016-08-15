using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace DraftKingsLineupGenerator
{
    public class NBAPlayerMatrix
    {
        private List<List<Player>> _allPlayers = new List<List<Player>> { }; //returning this
        private string _theFilePath;
        private List<Player> _pgs = new List<Player> { };
        private List<Player> _sgs = new List<Player> { };
        private List<Player> _sfs = new List<Player> { };
        private List<Player> _pfs = new List<Player> { };
        private List<Player> _cs = new List<Player> { };
        private List<Player> _gs = new List<Player> { };
        private List<Player> _fs = new List<Player> { };
        private List<Player> _utils = new List<Player> { };
        private int _pgCutoffCost;
        private int _sgCutoffCost;
        private int _sfCutoffCost;
        private int _pfCutoffCost;
        private int _cCutoffCost;
        private Excel.Application xlApp;
        private Excel.Workbook xlWorkbook;
        private Excel._Worksheet xlWorksheet;
        private Excel.Range xlRange;
        private int rowCount;
        private Player player;

        //Methods to read the positions, Names, Salaries of each player and add to respective lists
        public List<List<Player>> BuildPlayerList(int pgCutoff, int sgCutoff, int sfCutoff, int pfCutoff, int cCutoff, string fileNameHere)
        {
            this._pgCutoffCost = pgCutoff;
            this._sgCutoffCost = sgCutoff;
            this._sfCutoffCost = sfCutoff;
            this._pfCutoffCost = pfCutoff;
            this._cCutoffCost = cCutoff;
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

                GenerateNBALists(player); //Genertate Lists
            }

            GenerateNBAMatrix(); //Generate Matrix
        }

        private void GenerateNBALists(Player player)
        {
            //Select case to add players to lists
            switch (player.Position.ToString())
            {
                case "PG":
                    if (player.Cost >= _pgCutoffCost)
                    {
                        _pgs.Add(player);
                        _gs.Add(player);
                        _utils.Add(player);
                    }
                    break;
                case "SG":
                    if (player.Cost >= _sgCutoffCost)
                    {
                        _sgs.Add(player);
                        _gs.Add(player);
                        _utils.Add(player);
                    }
                    break;
                case "SF":
                    if (player.Cost >= _sfCutoffCost)
                    {
                        _sfs.Add(player);
                        _fs.Add(player);
                        _utils.Add(player);
                    }
                    break;
                case "PF":
                    if (player.Cost >= _pfCutoffCost)
                    {
                        _pfs.Add(player);
                        _fs.Add(player);
                        _utils.Add(player);
                    }
                    break;
                case "C":
                    if (player.Cost >= _cCutoffCost)
                    {
                        _cs.Add(player);
                        _utils.Add(player);
                    }
                    break;
            }
        }

        private void GenerateNBAMatrix()
        {
            //Add lists to matrix here...
            _allPlayers.Add(_pgs);
            _allPlayers.Add(_sgs);
            _allPlayers.Add(_sfs);
            _allPlayers.Add(_pfs);
            _allPlayers.Add(_cs);
            _allPlayers.Add(_gs);
            _allPlayers.Add(_fs);
            _allPlayers.Add(_utils);
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
