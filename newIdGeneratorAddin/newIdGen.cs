using ExcelDna.Integration;

namespace newIdGeneratorAddin
{
    public class newIdGenerator : IExcelAddIn
    {
        public void AutoOpen() { }
        public void AutoClose() { }

        [ExcelFunction(Description = "Converts a 15-digit ID into an 18-digit ID based on the Salesforce.com algorithm.",
                Category = "VLOOKUP Case Insensitivity")]
        public static string newId([ExcelArgument(Description = "15-digit ID that is ONLY 0-9 and A-Z to be converted to 18 digit unique ID.", Name = "originalId")]string inputValue)
        {
            try
            {
                //Input:    15-digit alphanumeric string
                //Output:   18-digit alphanumeric string whose last 3 are based on the first 3 groups of 5 characters

                //Step 1: Split inputValue into three sections
                char[][] charGroups = new char[3][];
                for (int i = 0; i < 3; i++)
                {
                    charGroups[i] = inputValue.ToCharArray(5 * i, 5);
                }

                //Step 2: Reverse the string the object-oriented way
                char[][] reversedChars = new char[3][];
                for (int i = 0; i < 3; i++)
                {
                    reversedChars[i] = new char[5];
                    for (int j = 0; j < 5; j++)
                    {
                        reversedChars[i][j] = charGroups[i][4 - j];
                    }
                }

                //Step 3/4: Lazy Lookup table.
                char[] algorithmAlphabet = new char[]{
	            'A','B','C','D','E','F','G','H','I','J','K','L',
                'M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z',
                '0','1','2','3','4','5','6','7','8','9' };

                string[] algorithmBinary = new string[]{
	            "00000","00001","00010","00011","00100","00101","00110","00111","01000","01001","01010","01011",
                "01100","01101","01110","01111","10000","10001","10010","10011","10100","10101","10110","10111",
                "11000","11001","11010","11011","11100","11101","11110","11111" };

                //Step 3: If the character is an Uppercase A-Z, then set it to 1. Otherwise, 0
                char[][] binaryChars = new char[3][];
                for (int k = 0; k < 3; k++)
                {
                    binaryChars[k] = new char[5];
                    for (int j = 0; j < 5; j++)
                    {
                        bool flag = true;
                        for (int i = 0; i < 26; i++)
                        {
                            if (reversedChars[k][j] == algorithmAlphabet[i])
                            {
                                flag = false;
                            }
                        }
                        binaryChars[k][j] = (flag) ? '0' : '1';
                    }
                }

                //Step 4: Do the binary replacements.
                string[] outputValues = new string[3];
                string returnValue = inputValue;
                for (int i = 0; i < 3; i++)
                {
                    outputValues[i] = new string(binaryChars[i]);
                }

                for (int k = 0; k < 3; k++)
                {
                    for (int i = 0; i < 32; i++)
                    {
                        if (outputValues[k] == algorithmBinary[i])
                        {
                            outputValues[k] = algorithmAlphabet[i].ToString();
                        }
                    }
                    returnValue += outputValues[k];
                }
                return returnValue;
            }
            catch
            {
                return "newId() error";
            }
        }
    }
}
