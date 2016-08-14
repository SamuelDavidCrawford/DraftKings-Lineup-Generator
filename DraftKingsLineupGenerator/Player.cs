using System;

namespace DraftKingsLineupGenerator
{
    public class Player
    {
        private string _name;
        private string _position;
        private int _cost;
        private int _iD;

        public string Name
        {
            get { return _name; }
            set {
                if (string.IsNullOrWhiteSpace(value))
                    throw new ArgumentNullException();
                _name = value; }
        }

        public string Position
        {
            get { return _position; }
            set {
                if (string.IsNullOrWhiteSpace(value))
                    throw new ArgumentNullException();
                _position = value; }
        }

        public int Cost
        {
            get { return _cost; }
            set {
                if (value == 0)
                    throw new ArgumentNullException();
                _cost = value; }
        }

        public int ID
        {
            get { return _iD; }
            set {
                if (value == 0)
                    throw new ArgumentNullException();
                _iD = value; }
        }
    }
}
