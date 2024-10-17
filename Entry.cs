using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kwerendy
{
    public class Entry
    {
        private string sign;
        public string Sign
        {
            get => sign;
            set { sign = value; }
        }

        private string date;
        public string Date
        {
            get => date;
            set { Date = value; }
        }

        private string action;
        public string Action
        {
            get => action;
            set { action = value; }
        }

        private string description;
        public string Description
        {
            get => description;
            set { description = value; }
        }

        private string name;
        public string Name
        {
            get => name;
            set { name = value; }
        }

        private string adressL1;
        public string AdressL1
        {
            get => adressL1;
            set { adressL1 = value; }
        }
        private string adressL2;
        public string AdressL2
        {
            get => adressL2;
            set { adressL2 = value; }
        }

        public string workplace;
        public string Workplace
        {
            get => workplace;
            set { workplace = value; }
        }

        private string initials;
        public string Initials
        {
            get => initials;
            set { initials = value; }
        }

        public Entry(string sign, string date, string action, string description, string name,
            string adressL1, string adressL2, string workplace, string initials)
        {
            this.sign = sign;
            this.date = date;
            this.action = action;
            this.description = description;
            this.name = name;
            this.adressL1 = adressL1;
            this.adressL2 = adressL2;
            this.workplace = workplace;
            this.initials = initials;
        }
    }
}
