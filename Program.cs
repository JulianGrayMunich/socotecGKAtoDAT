using System.Configuration;
using System.Globalization;
using System.Text;

using GNAgeneraltools;

//-- Purpose of Software
//-- To convert the Settop M1 gka file into an ASCII file with a user defined format for importation into 3rd party software
//-- The dat format is defined by Socotec
//-- gka file is in gons

//======[ Supression ]===================================


#pragma warning disable CS0164
#pragma warning disable CS8600
#pragma warning disable CS8601
#pragma warning disable CS8604
#pragma warning disable IDE0059


namespace socotecGKAtoDAT
{
    class Program
    {
        public static void Main()
        {


            // =====[ Library ]======================================

            gnaTools gnaT = new();

            //==== Console settings
            Console.OutputEncoding = System.Text.Encoding.Unicode;
            CultureInfo culture;
            culture = CultureInfo.CreateSpecificCulture("en-GB");

            // =====[ Configuration variables ]======================================

            string strFreezeScreen = ConfigurationManager.AppSettings["FreezeScreen"];
            string strProjectTitle = ConfigurationManager.AppSettings["ProjectTitle"];

            string strNoOfSettopFolders = ConfigurationManager.AppSettings["NoOfSettopFolders"];
            string strSettopFolder01 = ConfigurationManager.AppSettings["SettopFolder01"];
            string strSettopFolder02 = ConfigurationManager.AppSettings["SettopFolder02"];
            string strSettopFolder03 = ConfigurationManager.AppSettings["SettopFolder03"];
            string strSettopFolder04 = ConfigurationManager.AppSettings["SettopFolder04"];
            string strSettopFolder05 = ConfigurationManager.AppSettings["SettopFolder05"];
            string strDestinationPath = ConfigurationManager.AppSettings["DestinationPath"];
            string strTimekeeperPath = ConfigurationManager.AppSettings["TimekeeperPath"];
            string strSourceFileExtension = ConfigurationManager.AppSettings["SourceFileExtension"];
            string strOutputFileExtension = ConfigurationManager.AppSettings["OutputFileExtension"];
            string strCSVseparator = ConfigurationManager.AppSettings["CSVseparator"];
            string strDegRadGon = ConfigurationManager.AppSettings["DegRadGon"];
            string strSendEmails = ConfigurationManager.AppSettings["SendEmails"];
            string strEmailLogin = ConfigurationManager.AppSettings["EmailLogin"];
            string strEmailPassword = ConfigurationManager.AppSettings["EmailPassword"];
            string strEmailFrom = ConfigurationManager.AppSettings["EmailFrom"];
            string strNoDataAlarmRecipients = ConfigurationManager.AppSettings["NoDataAlarmRecipients"];
            string strNoDataInterval = ConfigurationManager.AppSettings["NoDataInterval_hrs"];
            string testEmail = ConfigurationManager.AppSettings["testEmail"];
            string timeToSendSystemOperationalEmail = ConfigurationManager.AppSettings["SystemOperationalEmail"];

            // =====[ Define variables ]======================================

            string strCurrentGKAfolder = " ";
            string strLastProcessedGKAFile = "";
            string strLastProcessedGKAFileName = "";
            string strLastProcessedGKAFolder = "";
            string textFile;
            string text;
            string strTimekeeperFile;
            string strCurrentGKAfile = "";
            string[] SettopFolder = new string[10];
            string[] GKAfolder = new string[300];

            string strActiveGKAfolder = "";

            string[] element = new string[35];

            //-- Variable Declerations
            string strSourceFile, strOutputFile;
            string strLineHolder;
            string strConvertedLine;
            string strATS, strTarget;
            string strATSserialNumber;
            string mnDateTime, strTimeStamp;
            string strRecordCounter, strObsCounter, strFace, strHA, strVA, strSD, strPsmCnst;
            string specifier;
            string strAnswer;
            string strFace1 = "blank";
            string strFace2 = "blank"; ;
            string strTarget1, strFaceLeft, strTarget2, strFaceRight;
            string strMeanHA, strMeanVA, strMeanSD, strNoise;
            string strConvertedDataLine = "";
            string strRecordNumber = "RecordNumber";
            string strFaceFlag = "F1F2";
            string strDateTimeofThisRound;
            string strDataForTimekeeper = "ATS;TimeStamp;OK;0";
            string strTemperature = "";
            string strPressure = "";
            string strPrismConstant = "";
            string strFirstDataLineFlag = "Yes";

            int intNoOfLines;
            int gpsWk;
            int intRecordNumber;
            int intObsCounter, intRecordCounter, intConvertedObsCounter;
            int i, j, k, m;
            int counter;
            int iNoOfSettopFolders;

            double gpsSc;
            double dblMeanHA, dblMeanVA;
            double dblConversionFactor, dblConvertedData;

            string[] strSettopDataFiles = new string[1000];
            string[] lines = new string[10000];
            string[] element1 = new string[35];
            string[] element2 = new string[35];
            string[] element3 = new string[35];
            string[] element4 = new string[5];
            string[] Timekeeper = new string[5];
            string[] CurrentLine = new string[5];
            string[] Round = new string[2000];
            string[] MeanRound = new string[1000];
            string[] HeaderLines = new string[5];

            // assign variables
            intObsCounter = 99;
            intRecordCounter = 10000;
            intRecordNumber = 1;
            intConvertedObsCounter = -1;

            dblConversionFactor = 1.0;
            dblConvertedData = 0.0;
            specifier = "G";
            strOutputFile = "";
            strMeanHA = "";
            strMeanVA = "";
            strTimeStamp = "";
            string strGKAfile = "";
            string strSettopM1 = "";

            //==== [Procedures]===========================================================================================

            string[] createHeaderLines(string strDateTimeStamp, string strRTSname, string strTemperatureValue, string strPressureValue)
            {
                String[] strHeaderLines = new String[5];
                strHeaderLines[0] = strDateTimeStamp + "," + "10001" + ",\"" + strRTSname + "\",\"Lgr_Environ\",10.00,20.00,30.00,4000,50";
                strHeaderLines[1] = strDateTimeStamp + "," + "10002" + ",\"" + strRTSname + "\",\"RTS_Environ\",90,9.3,1,4000,50";
                strHeaderLines[2] = strDateTimeStamp + "," + "10003" + ",\"" + strRTSname + "\",\"RTS_Corrections\"" + "," + strTemperatureValue + "," + strPressureValue + ",5,4000,50";
                strHeaderLines[3] = strDateTimeStamp + "," + "10004" + ",\"" + strRTSname + "\",\"RTS_Compensator\",0.0,0.0,0.0,4000,50";
                strHeaderLines[4] = strDateTimeStamp + "," + "10005" + ",\"" + strRTSname + "\",\"RTS_CTAngle\",0.0,0.0,0.0,4000,50";
                return strHeaderLines;
            }

            string ProcessGKAFile()
            {

                strSourceFile = strCurrentGKAfile;

                Console.WriteLine(strSourceFile);

                lines = File.ReadAllLines(strSourceFile);                     // the settop data is now in the array lines  
                intNoOfLines = lines.Length - 1;                              // first element starts at 0

                // prepare the format of the array lines[], doing text replacement

                i = 0;
                strATS = "";
                foreach (string line in lines)
                {
                    strLineHolder = line;
                    //reset the element1 array
                    for (k = 0; k == 35; k++)
                    {
                        element1[k] = "x";
                    }

                    // Now split the line into its components
                    element1 = strLineHolder.Split(',');

                    // Now make up the new line
                    strConvertedLine = "";
                    if (element1[0] == "#GNV11")
                    {
                        strConvertedLine = element1[0];
                        intRecordCounter = 0;
                        goto nextLine;
                    }
                    else if (element1[0] == "#END11")
                    {
                        // All the observations for the set of rounds have been extracted and stored in Round[]
                        // range 0..intRecordCounter

                        intRecordCounter++;
                        strConvertedLine = element1[0];

                        //Sort to reflect F1/F2.. 
                        Array.Sort(Round, 0, intRecordCounter);
                        intConvertedObsCounter = -1;

                        CreateOutputFileName();

                        // Create the unique record number
                        double result = DateTime.Now.Subtract(DateTime.MinValue).TotalSeconds - 63700000000.00;
                        intRecordNumber = Convert.ToInt32(result);

                        setFaceFlag();

                        // **** checked - OK

                        // walk through the observations of the round and create the meaned observation for each target
                        for (j = 1; j < intRecordCounter - 1; j++)
                        {
                            strFace1 = Round[j];
                            strFace2 = Round[j + 1];
                            element1 = strFace1.Split(',');
                            element2 = strFace2.Split(',');
                            strTarget1 = element1[0];
                            strTarget2 = element2[0];
                            string strDataTime = element1[2];   // bad programming
                            strTarget = element1[4];
                            strPsmCnst = element1[8];
                            strNoise = element1[9];
                            strTarget1 = element1[0];
                            strFaceLeft = element1[1];

                            strTarget2 = element2[0];
                            strFaceRight = element2[1];

                            if ((strTarget1 == strTarget2) && (strFaceLeft == "1") && (strFaceRight == "2") && (strFaceFlag == "F1F2"))
                            {
                                // the two lines represent Face 1 and face 2 observations onto the same target and the data can be meaned.

                                string mnTarget = element1[0];  //bad programming
                                mnDateTime = element1[2];

                                string mnPrismConstant = element1[8];  //bad programming
                                strNoise = element1[9];
                                strRecordNumber = Convert.ToString(intRecordNumber++);

                                computeMeanAngles();

                                // compute mean distance
                                strMeanSD = Convert.ToString((Convert.ToDouble(element1[7]) + Convert.ToDouble(element2[7])) / 2.0);

                                // if the distance is not zero, create a meaned observation line and append to MeanRound
                                if ((Convert.ToDouble(element1[7]) > 0.01) && (Convert.ToDouble(element2[7]) > 0.01))
                                {
                                    createOutputDataLine();
                                    intConvertedObsCounter++;
                                    MeanRound[intConvertedObsCounter] = strConvertedDataLine;
                                }

                                //Console.WriteLine(strATS);
                                //Console.WriteLine("Face 1: "+strFace1);
                                //Console.WriteLine("Face 2: " + strFace2);
                                //Console.WriteLine("a  "+ strConvertedDataLine);
                            }

                            if ((strFaceLeft == "1") && (strFaceFlag == "F1"))
                            {
                                // this is a face 1 only set of observations.

                                string mnTarget = element1[0];  //bad programming
                                mnDateTime = element1[2];

                                string mnPrismConstant = element1[8];  //bad programming
                                strNoise = element1[9];
                                strRecordNumber = Convert.ToString(intRecordNumber++);

                                computeMeanAngles();
                                // compute mean distance
                                strMeanSD = Convert.ToString(Convert.ToDouble(element1[7]));

                                // if the distance is not zero, create an observation line and append to MeanRound
                                if (Convert.ToDouble(element1[7]) > 0.01)
                                {
                                    createOutputDataLine();
                                    intConvertedObsCounter++;
                                    MeanRound[intConvertedObsCounter] = strConvertedDataLine;
                                }

                            }

                        }


                        if (intConvertedObsCounter > 2)
                        {
                            writeConvertedObsToFile();
                        }

                    }
                    else if ((element1[13] == "1") & (element1[14] == "100"))
                    {
                        strATS = element1[1];
                        strATSserialNumber = element1[0];
                        goto nextLine;
                    }
                    else
                    {
                        processSettopObservationLine();
                    }

nextLine:
                    continue;

                }

                strAnswer = "Process GKA File";
                return strAnswer;
            }

            string UpdateDataTimeStamp()
            {
                // to write the ATS and associated latest observation timestamp to a file
                // The file is stored in the Timekeeper folder
                // This file is used to detect a data break and associated data alarm

                // Locate the array element that has the same name as the ATS
                // Replace the time stamp component
                // If the ATS was not found, then append to the end of the array
                // Delete the time stamp file
                // Recreate the time stamp file
                // Write the array to the time stamp file

                string strTimeKeeperFlag = "No";

                Console.WriteLine(strDataForTimekeeper);

                CurrentLine = strDataForTimekeeper.Split(';');
                string currentATS = CurrentLine[0];
                string currentTimestamp = CurrentLine[1];
                string currentAlarmState = CurrentLine[2];
                string currentHours = CurrentLine[3];


                strTimekeeperFile = strTimekeeperPath + "Timekeeper.txt";
                if (!File.Exists(strTimekeeperFile))
                {
                    using StreamWriter writetext = new(strTimekeeperFile, false);
                    writetext.WriteLine(strDataForTimekeeper);
                    writetext.Close();
                }
                else
                {
                    //Read the contents of the Timekeeper file into an array
                    string[] TimeLine = File.ReadAllLines(strTimekeeperFile);
                    // step through the array & split the elements
                    int iTimeLineCounter = TimeLine.Length;
                    for (m = 0; m < iTimeLineCounter; m++)
                    {
                        string strTemp = TimeLine[m];
                        Timekeeper = TimeLine[m].Split(';');

                        string TimekeeperATS = Timekeeper[0];
                        string TimekeeperTimestamp = Timekeeper[1];
                        string TimekeeperAlarmState = Timekeeper[2];
                        string TimekeeperHours = Timekeeper[3];

                        if (currentATS == TimekeeperATS)
                        {
                            // Update the time stamp
                            int comparison = String.Compare(currentTimestamp, TimekeeperTimestamp, comparisonType: StringComparison.OrdinalIgnoreCase);
                            if (comparison > 0)
                            {
                                TimekeeperTimestamp = currentTimestamp;
                                TimeLine[m] = TimekeeperATS + ";" + TimekeeperTimestamp + ";" + "OK;updated";
                                // Write the array with the updated timestamp,alarm state & hours to the Timekeeper file
                                File.WriteAllLines(strTimekeeperFile, TimeLine, Encoding.UTF8);
                                strTimeKeeperFlag = "Updated";
                                goto ExitRoutine;
                            }
                        }
                    }
                }

                // if the strDataForTimekeeper != "ATS;TimeStamp;OK;0" then do this...


                if (strTimeKeeperFlag == "No")
                {
                    //append the timestamp data strDataForTimekeeper
                    File.AppendAllText(strTimekeeperFile, strDataForTimekeeper);
                }

ExitRoutine:
                strAnswer = "UpdateDataTimeStamp";
                return strAnswer;
            }

            string setFaceFlag()
            {

                // step through all the rounds to find a F2 observation => F1/F2
                // If no F2 observations are found then it is a F1 set
                for (j = 1; j < intRecordCounter - 1; j++)
                {
                    strFaceFlag = "F1";
                    strFace1 = Round[j];
                    element1 = strFace1.Split(',');
                    strFace = element1[1];

                    if (strFace == "2")
                    {
                        strFaceFlag = "F1F2";
                        goto label1;
                    }
                }
label1:
                strAnswer = "setFaceFlag";
                return strAnswer;
            }

            string processSettopObservationLine()
            {

                intRecordCounter++;
                strRecordCounter = Convert.ToString(intRecordCounter);

                //generate the date/time of the observation;
                // week 0 : 23:59:42 UTC on 6 April 2019

                // Extract data
                strTemperature = element1[25];
                strPressure = element1[26];
                strPrismConstant = element1[15];

                gpsWk = Convert.ToInt16(element1[2]);
                gpsSc = Convert.ToDouble(element1[4]);
                strTimeStamp = GetFromGPS(gpsWk, gpsSc);
                strDateTimeofThisRound = strTimeStamp;
                //checkTheProcessRoundFlag();


                // Check here whether this set is later than the last set that was processed
                // If no, then skip to the next set of observations

                // Extract the face and Target information
                strFace = element1[6];
                strTarget = element1[1];

                // Convert the Ha and Va from gons to the selected value (normally decimal degrees)
                // The selected output units are defibed in the system variable strDegRadGon
                // This determines the dblConversionFactor

                setConversionFactor();

                dblConvertedData = Convert.ToDouble(element1[10]) * dblConversionFactor;
                strHA = dblConvertedData.ToString(specifier, CultureInfo.InvariantCulture);
                dblConvertedData = Convert.ToDouble(element1[12]) * dblConversionFactor;
                strVA = dblConvertedData.ToString(specifier, CultureInfo.InvariantCulture);
                strSD = element1[7];
                strPsmCnst = element1[15];
                strObsCounter = Convert.ToString(intObsCounter);

                //Create the new line
                strConvertedLine = strATS + "," + strTarget + "," + strHA + "," + strVA + "," + strSD + "," + strPsmCnst + ",33423511";
                strConvertedLine = strTarget + "," + strFace + "," + strTimeStamp + "," + strRecordCounter + ":" + strConvertedLine;

                if (strFirstDataLineFlag == "Yes")
                {
                    HeaderLines = createHeaderLines(strTimeStamp, strATS, strTemperature, strPressure);
                    strFirstDataLineFlag = "No";
                }

                intObsCounter++;
                Round[intRecordCounter] = strConvertedLine;

                strAnswer = "processSettopObservationLine";
                return strAnswer;
            }

            string writeConvertedObsToFile()
            {
                strOutputFile = strDestinationPath + strOutputFile;
                // Write the meanRound to the strOutputFile file

                if (!File.Exists(strOutputFile))
                {
                    using StreamWriter writetext = new(strOutputFile, false);

                    for (j = 0; j <= 4; j++)
                    {
                        writetext.WriteLine(HeaderLines[j]);

                    }


                    for (j = 0; j <= intConvertedObsCounter; j++)
                    {
                        writetext.WriteLine(MeanRound[j]);
                    }



                    writetext.Close();
                }

                intConvertedObsCounter = -1;
                strAnswer = "writeConvertedObstoFile";
                return strAnswer;
            }

            string createOutputDataLine()
            {
                string Field1, Field2, Field3, Field4, Field5, Field6, Field7, Field8, Field9;

                Field1 = mnDateTime;
                Field2 = strRecordNumber;
                Field3 = "\"" + strATS + "\"";
                Field4 = "\"" + strTarget + "\"";
                Field5 = strMeanHA;
                Field6 = strMeanVA;
                Field7 = strMeanSD;
                Field8 = strPsmCnst;
                Field9 = strNoise;
                strConvertedDataLine = Field1 + strCSVseparator + Field2 + strCSVseparator + Field3 + strCSVseparator + Field4 + strCSVseparator + Field5 + strCSVseparator + Field6 + strCSVseparator + Field7 + strCSVseparator + Field8 + strCSVseparator + Field9;
                strAnswer = "computeOutputDataLine";
                return strAnswer;
            }

            string CreateOutputFileName()
            {
                element = strCurrentGKAfile.Split('\\');
                var count = strCurrentGKAfile.Split('\\').Length;
                element = element[count - 1].Split('.');
                strOutputFile = strATS + "_" + element[0] + "." + strOutputFileExtension;
                strAnswer = "createOutputFileName";
                return strAnswer;
            }

            string computeMeanAngles()
            {
                if (strFaceFlag == "F1F2")
                {
                    // mean HA
                    double dblFace1 = Convert.ToDouble(element1[5]);
                    double dblFace2 = Convert.ToDouble(element2[5]);

                    dblFace2 = (dblFace2 >= 0.0) && (dblFace2 < 180.0) ? dblFace2 + 180.0 : dblFace2 - 180.0;

                    dblMeanHA = (dblFace1 + dblFace2) / 2.0;
                    strMeanHA = Convert.ToString(dblMeanHA);

                    // mean VA
                    dblFace1 = Convert.ToDouble(element1[6]);
                    dblFace2 = Convert.ToDouble(element2[6]);

                    dblMeanVA = dblFace1 < dblFace2 ? dblFace1 : dblFace2;

                    dblMeanVA += (360.0 - dblFace1 - dblFace2) / 2.0;
                    strMeanVA = Convert.ToString(dblMeanVA);

                }
                if (strFaceFlag == "F1")
                {
                    // mean HA
                    double dblFace1 = Convert.ToDouble(element1[5]);
                    dblMeanHA = dblFace1;
                    strMeanHA = Convert.ToString(dblMeanHA);
                    // mean VA
                    dblFace1 = Convert.ToDouble(element1[6]);
                    dblMeanVA = dblFace1;
                    strMeanVA = Convert.ToString(dblMeanVA);
                }

                strAnswer = "computeMeanAngles";
                return strAnswer;
            }

            string GetFromGPS(int weeknumber, double seconds)
            {
                DateTime datum = new(1980, 1, 6, 0, 0, 0);
                DateTime week = datum.AddDays(weeknumber * 7);
                DateTime time = week.AddSeconds(seconds);
                string strDT = time.ToString("s");
                string OldValue = "T";
                string NewValue = " ";
                strDT = strDT.Replace(OldValue, NewValue);
                strDT = "\"" + strDT + "\"";
                return strDT;
            }

            string setConversionFactor()
            {

                // set conversion factors
                // Source data is in gons
                dblConversionFactor = strDegRadGon switch
                {
                    "Deg" => 180.0 / 200.0,
                    "Rad" => 3.14159265358979 / 200.0,
                    "Gon" => 1.0,
                    _ => 0.0,
                };
                strAnswer = "setConversionFactor";
                return strAnswer;

            }

            // =====[Main Program] ======================================

            gnaT.WelcomeMessage("socotecGKAtoDAT 20230630");
            string strSoftwareLicenseTag = "GKADAT";
            _ = gnaT.checkLicenseValidity(strSoftwareLicenseTag, strProjectTitle, strEmailLogin, strEmailPassword, strSendEmails);


            // allocate the Settop folders to a SettopFolder array
            iNoOfSettopFolders = Convert.ToInt16(strNoOfSettopFolders);
            for (i = 1; i <= iNoOfSettopFolders; i++)
            {
                if (i == 1) { SettopFolder[i] = strSettopFolder01; };
                if (i == 2) { SettopFolder[i] = strSettopFolder02; };
                if (i == 3) { SettopFolder[i] = strSettopFolder03; };
                if (i == 4) { SettopFolder[i] = strSettopFolder04; };
                if (i == 5) { SettopFolder[i] = strSettopFolder05; };
            }

            Array.Sort(SettopFolder, 1, iNoOfSettopFolders);

            // Loop through the Settop folders
            for (counter = 1; counter <= iNoOfSettopFolders; counter++)
            {
                strCurrentGKAfolder = SettopFolder[counter];

                Console.WriteLine("=====[ " + strCurrentGKAfolder + " ]==========================================");

                // Get the name of the last GKA file; that was processed
                textFile = strCurrentGKAfolder + "latestProcessedGKAfile.txt";

                if (!File.Exists(textFile))
                {
                    using StreamWriter writetext = new(textFile, false);
                    writetext.WriteLine(strCurrentGKAfolder + "20000101" + "\\" + "20000101_000001_20000101_000001.gka" + "\\end");
                    writetext.Close();
                }

                text = File.ReadAllText(textFile);
                // split this string to extract the name of the last folder that contains the file that was processed
                element = text.Split('\\');
                var count = text.Split('\\').Length;
                strLastProcessedGKAFileName = element[count - 2];
                strSettopM1 = element[count - 3];
                strGKAfile = element[count - 1];
                strLastProcessedGKAFolder = "";
                for (j = 0; j < count - 2; j++)
                {
                    strLastProcessedGKAFolder = strLastProcessedGKAFolder + element[j] + "\\";
                }
                strLastProcessedGKAFile = text;

                // Get all subdirectories
                string[] subdirectoryEntries = Directory.GetDirectories(strCurrentGKAfolder);
                // Now step through each subdirectory, locate the gka files and process
                foreach (string subdirectory in subdirectoryEntries)
                {
                    string strSubDirectory = subdirectory;
                    string strProcessFolderFlag = "No";
                    if (String.Compare(strSubDirectory, strActiveGKAfolder) >= 0) { strProcessFolderFlag = "Yes"; };

                    if (strProcessFolderFlag == "Yes")
                    {
                        // this folder contains gka files that must be processed
                        strSubDirectory += "\\";

                        // Now step through the gka files in this folder
                        string[] fileArray = Directory.GetFiles(strSubDirectory);
                        foreach (string gkaFile in fileArray)
                        {

                            strCurrentGKAfile = gkaFile;

                            Console.WriteLine(strCurrentGKAfile);

                            // Check whether this gka file must be processed or not.
                            string strProcessFileFlag = "No";
                            if (String.Compare(strCurrentGKAfile, strLastProcessedGKAFile) > 0) { strProcessFileFlag = "Yes"; };

                            if (strProcessFileFlag == "Yes")
                            {
                                ProcessGKAFile();
                                // the gka file has been processed and we have the latest observation data for the Timekeeper folder
                                strDataForTimekeeper = strATS + ";" + strTimeStamp + ";OK;0";
                            }
                        }
                    }
                }

                // Update the latestProcessedGKAfile file
                File.Delete(textFile);
                using (StreamWriter writetext = new(textFile, false))
                {
                    writetext.WriteLine(strCurrentGKAfile);
                    writetext.Close();
                }

                UpdateDataTimeStamp();

                // At this stage, check the data time interval for no data
                Console.WriteLine("No data check: " + strCurrentGKAfolder);
                //gnaT.noDataAlarm(strCurrentGKAfolder, strNoDataInterval, strEmailLogin, strEmailPassword, strEmailFrom, strNoDataAlarmRecipients, testEmail);

            }

ThatsAllFolks:

            gnaT.freezeScreen(strFreezeScreen);

            Console.WriteLine("\nProcessing completedd...");
            Environment.Exit(0);

        }
    }
}