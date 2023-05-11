using System.Linq;

namespace BlockInsertionV3
{
    class ID_Reader
    {
        //need to fine tune use

        public static string interpreter(string styleID)
        {
            string styleIDName = "Undefined";
            //check if there is a styleID ie not null and not blank
            if (styleID != null && styleID != "")
            {
                //split the styleID up by hyphens *typical to naming convention
                //if there are no hyphens it is likely a sleeper unit
                string[] styleParts = styleID.Split('-');
                if (styleParts.Count() > 1)
                {
                    //ignore the first part for now, group number isnt relevant
                    //examine the second and third and possibly fourth parts
                    //more parts than this are counter to convention and cannot be accounted for right now
                    string example = styleParts[1];
                    string part = example.Substring(0, 2);
                    part = getModel(part);

                    string part2 = example.Substring(2, 2);
                    part2 = getDetails(part2);

                    //Piece together parts for new name
                    styleIDName = part2 + part;

                    if (example.Length > 4)
                    {
                        //Piece together parts for new name
                        string designater = example.Substring(3, 1);
                        if (designater == "L")
                        {
                            //Piece together parts for new name
                            styleIDName = "Left" + part2 + part;
                        }
                        else if (designater == "R")
                        {
                            //Piece together parts for new name
                            styleIDName = "Right" + part2 + part;
                        }
                    }

                    if (styleParts.Count() > 2)
                    {
                        string last = styleParts[2];
                        last = finalPart(last);
                        //Piece together parts for new name
                        styleIDName = last + part2 + part;
                    }
                }
                else
                { styleIDName = "Sleeper"; }
            }
            return styleIDName;
        }

        //decide if the last parts are a side or number or such
        private static string finalPart(string lastPart)
        {
            string returnString;
            switch (lastPart)
            {
                case "X":
                    returnString = "";
                    break;

                case "L":
                    returnString = "Left";
                    break;

                case "R":
                    returnString = "Right";
                    break;

                default:
                    returnString = lastPart + '"';
                    break;
            }
            return returnString;
        }

        //compare the second two digits to naming chart
        private static string getDetails(string detString)
        {
            string returnString;
            switch (detString)
            {
                case "00":
                    returnString = "";
                    break;

                case "02":
                    returnString = "ARM";
                    break;

                case "05":
                    returnString = "SWIVEL";
                    break;

                case "06":
                    returnString = "SWIVEL/GLIDER";
                    break;

                case "07":
                    returnString = "GLIDER";
                    break;

                case "08":
                    returnString = "ROCKER";
                    break;

                case "09":
                    returnString = "SWIVEL/ROCKER";
                    break;

                case "10":
                    returnString = "DAYBED";
                    break;

                case "15":
                    returnString = "CHAIR 1/2";
                    break;

                case "17":
                    returnString = "TETE A TETE";
                    break;

                case "20":
                    returnString = "CHAISE";
                    break;

                case "21":
                    returnString = "ARM CHAISE";
                    break;

                case "25":
                    returnString = "ARMLESS CHAISE";
                    break;

                case "26":
                    returnString = "ANGLED CHAISE";
                    break;

                case "31":
                    returnString = "ARMLESS CHAIR";
                    break;

                case "32":
                    returnString = "ARMLESS LOVESEAT";
                    break;

                case "33":
                    returnString = "ARMLESS SOFA";
                    break;

                case "41":
                    returnString = "ARM CHAIR";
                    break;

                case "42":
                    returnString = "ARM LOVESEAT";
                    break;

                case "43":
                    returnString = "ARM SOFA";
                    break;

                case "51":
                case "52":
                case "53":
                    returnString = "BUMPER";
                    break;

                case "61":
                    returnString = "CORNER CHAIR";
                    break;

                case "62":
                    returnString = "CORNER LOVESEAT";
                    break;

                case "63":
                    returnString = "CORNER SOFA";
                    break;

                case "64":
                    returnString = "CURVE";
                    break;

                default:
                    returnString = "";
                    break;
            }
            return returnString;
        }

        //compares the first two digits to a naming chart
        private static string getModel(string modString)
        {
            string returnString;
            switch (modString)
            {
                case "10":
                case "11":
                case "12":
                case "13":
                    returnString = "DINING";
                    break;
                case "15":
                    returnString = "BAR STOOL";
                    break;

                case "20":
                case "21":
                case "22":
                case "23":
                case "24":
                    returnString = "CHAIR";
                    break;

                case "30":
                case "31":
                case "32":
                case "33":
                case "34":
                    returnString = "OTTOMAN";
                    break;

                case "40":
                case "41":
                case "42":
                case "43":
                case "44":
                    returnString = "LOVESEAT";
                    break;

                case "50":
                case "51":
                case "52":
                case "53":
                case "54":
                    returnString = "SOFA";
                    break;

                case "60":
                case "61":
                case "62":
                case "63":
                case "64":
                    returnString = "BENCH";
                    break;

                case "70":
                    returnString = "BED";
                    break;

                case "80":
                case "82":
                case "84":
                    returnString = " ";
                    break;

                default:
                    returnString = "ERROR";
                    break;
            }
            return returnString;
        }
    }
}
