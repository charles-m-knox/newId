using System;
using ExcelDna.Integration;

namespace newIdGeneratorAddin
{
    public class newIdGenerator : IExcelAddIn
    {
        public void AutoOpen() { }
        public void AutoClose() { }
        

        [ExcelFunction(Description = "Converts a 15-digit ID into an 18-digit ID based on the Salesforce.com algorithm.",
                Category = "VLOOKUP Case Insensitivity")]
        public static string newId([ExcelArgument(Description="15-digit ID that is ONLY 0-9 and A-Z to be converted to 18 digit unique ID.", Name="originalId")]string inputValue)
        {
            try
            {
                //Input:    15-digit alphanumeric string
                //Output:   18-digit alphanumeric string whose last 3 are based on the first 3 groups of 5 characters

                //Step 1: Split inputValue into three sections
                char[] charGroup1 = inputValue.ToCharArray(0, 5);
                char[] charGroup2 = inputValue.ToCharArray(5, 5);
                char[] charGroup3 = inputValue.ToCharArray(10, 5);

                //Step 2: Reverse the string the object-oriented way
                char[] revGroup1 = new char[5];
                char[] revGroup2 = new char[5];
                char[] revGroup3 = new char[5];
                for (int i = 0; i < 5; i++)
                {
                    revGroup1[i] = charGroup1[4 - i];
                    revGroup2[i] = charGroup2[4 - i];
                    revGroup3[i] = charGroup3[4 - i];
                }


                //Step 3/4: Lazy Lookup table.
                char[] algorithmAlphabet = new char[]{
	            'A','B','C','D','E','F','G','H','I','J','K','L',
                'M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z',
                '0','1','2','3','4','5','6','7','8','9'
            };
                string[] algorithmBinary = new string[]{
	            "00000","00001","00010","00011","00100","00101","00110","00111","01000","01001","01010","01011",
                "01100","01101","01110","01111","10000","10001","10010","10011","10100","10101","10110","10111",
                "11000","11001","11010","11011","11100","11101","11110","11111"
            };
                //Step 3: If the character is an Uppercase A-Z, then set it to 1. Otherwise, 0
                //revGroup1:
                for (int j = 0; j < 5; j++)
                {
                    bool flag = true;
                    for (int i = 0; i < 26; i++)
                    {
                        if (revGroup1[j] == algorithmAlphabet[i])
                        {
                            flag = false;
                        }
                    }
                    if (flag)
                    {
                        revGroup1[j] = '0';
                    }
                    else
                    {
                        revGroup1[j] = '1';
                    }
                }
                //revGroup2:
                for (int j = 0; j < 5; j++)
                {
                    bool flag = true;
                    for (int i = 0; i < 26; i++)
                    {
                        if (revGroup2[j] == algorithmAlphabet[i])
                        {
                            flag = false;
                        }
                    }
                    if (flag)
                    {
                        revGroup2[j] = '0';
                    }
                    else
                    {
                        revGroup2[j] = '1';
                    }
                }
                //revGroup3:
                for (int j = 0; j < 5; j++)
                {
                    bool flag = true;
                    for (int i = 0; i < 26; i++)
                    {
                        if (revGroup3[j] == algorithmAlphabet[i])
                        {
                            flag = false;
                        }
                    }
                    if (flag)
                    {
                        revGroup3[j] = '0';
                    }
                    else
                    {
                        revGroup3[j] = '1';
                    }
                }
                //Step 4: Do the replacements.
                //00000 = A, 00001 = B, 00010 = C, etc, until 11110 = 4, 11111 = 5
                string outputValue1 = new string(revGroup1), outputValue2 = new string(revGroup2), outputValue3 = new string(revGroup3);

                for (int i = 0; i < 32; i++)
                {
                    if (outputValue1 == algorithmBinary[i])
                    {
                        outputValue1 = algorithmAlphabet[i].ToString();
                    }
                    if (outputValue2 == algorithmBinary[i])
                    {
                        outputValue2 = algorithmAlphabet[i].ToString();
                    }
                    if (outputValue3 == algorithmBinary[i])
                    {
                        outputValue3 = algorithmAlphabet[i].ToString();
                    }
                }

                string outputValue = string.Concat(inputValue, outputValue1, outputValue2, outputValue3);

                return outputValue;
            }
            catch
            {
                return "newId() error";
            }
        }
    }
}
