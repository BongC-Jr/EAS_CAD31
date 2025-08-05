using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EASI_CAD31
{
    public class CCData
    {
        /**
         * Common Command Data
         * Essentially these are static properties
         */
        static string licenseCode;
        static Boolean isActive;
        static string licenseNotes;
        static string slabGSheetID;
        static string slabDesignData;
        static List<string> dimStyle;
        static string selectedDimstyle;
        static int numOfSelection; //SOR command

        public void setNumOfSelection(int numberOfSelection)
        {
            numOfSelection = numberOfSelection;
            return;
        }
        public int getNumOfSelection()
        {
            return numOfSelection;
        }
        public void setSelectedDimStyle(string dimStyle)
        {
            selectedDimstyle = dimStyle;
            return;
        }
        public string getSelectedDimStyle() { return selectedDimstyle; }

        public void setDimStyle(List<string> lsDimStyle)
        {
            dimStyle = lsDimStyle;
            return;
        }
        public List<string> getDimStyle()
        {
            return dimStyle;
        }

        public void setSlabDesignData(string strVal)
        {
            slabDesignData = strVal;
            return;
        }

        public string getSlabDesignData()
        {
            return slabDesignData;
        }

        public void setSlabGSheetID(string strVal)
        {
            slabGSheetID = strVal;
            return;
        }

        public string getSlabGSheetID()
        {
            return slabGSheetID;
        }

        public void setLicenseNote(string strVal)
        {
            licenseNotes = strVal;
            return;
        }

        public string getLicenseNote()
        {
            return licenseNotes;
        }

        public void setLicenseStatus(Boolean boolVal)
        {
            isActive = boolVal;
            return;
        }

        public Boolean isLicenseActive()
        {
            return isActive;
        }

        public void setLicense(string strVal)
        {
            licenseCode = strVal;
            return;
        }

        public string getLicense()
        {
            return licenseCode;
        }
    }
}
