#define USE_COMPANIES_IN_CONFIG //Comment out to use all companies in database
#define USE_RESERVES_IN_CONFIG //Comment out to use all reserves in database

using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Diagnostics;
using System.Windows.Forms;

namespace AppPivotNet
{
    class Program
    {

        static private PLEXOS7_NET.Core.DatabaseCore db;
        static private PLEXOS7_NET.Core.Solution zip;
        static private Model ModelBase; //Base model data from which all plexos models are derived
        static private String strInputFile; //Full path to plexos xml file
        static private String strOutputFile; //Full path to plexos zip output file

        //private static List<String> lconfig_reserves = new List<String>(); //list of reserve names as defined in config file
        private static List<String> lconfig_companies = new List<String>(); //list of company names as defined in config file
        private static List<String> lconfig_base_scenarios = new List<String>(); //list of basic scenario names as defined in config file
        private static List<String> lconfig_base_horizons = new List<String>(); //list of horizon names as defined in config file
        private static List<String> lconfig_models_new = new List<String>(); //list of model names for this pivot analysis
        private static List<ReportProperty> lconfig_base_reports = new List<ReportProperty>();

        //private const String OUT_PATH = "D:/01.pivotal/intermediate/Test_Model/Model PRGdia_Full_Definitivo Solution.zip"; //Debug only
        private const String RESERVE_CATEGORY = "RESERVE"; //Hard-wired name of category for reserve-related scenarios
        private const String CONFIG_FILE = "pivotal.json"; //Hard-wired name of config file
        private const String ASSEMBLY_FILE = "OpenIndices.dll"; //Hard-wired name of assembly for computing the static indices
        private const String ASSEMBLY_NAMESPACE = "OpenIndices";
        private const double MAX_RESPONSE_ZERO = 0.001; //This is another workaround for PLEXOS 8-1 because it seems it's not fully deleting provision variable when max response is zero

        #region Input Calculations

        private static void InitConnection()
        {
            db = new PLEXOS7_NET.Core.DatabaseCore();
            db.Connection(strInputFile);
        }

        private static void InitInput()
        {
            //Objects: System, Generators, Reserves, Companies
            String[] strSystems = db.GetObjects(EEUTILITY.Enums.ClassEnum.System);
            Elements.SystemName = strSystems[0];
            String[] strGens = db.GetObjects(EEUTILITY.Enums.ClassEnum.Generator);
            foreach (String g in strGens)
            {
                Elements.Add(EEUTILITY.Enums.ClassEnum.Generator, g);
            }
            List<String> strCompanies;
#if USE_COMPANIES_IN_CONFIG
            strCompanies = new List<String>(lconfig_companies);
#else
            strCompanies = new List<String>(db.GetObjects(EEUTILITY.Enums.ClassEnum.Company));
#endif
            foreach (String c in strCompanies)
            {
                Elements.Add(EEUTILITY.Enums.ClassEnum.Company, c);
            }
#if USE_RESERVES_IN_CONFIG

#else
            String[] strReserves = db.GetObjects(EEUTILITY.Enums.ClassEnum.Reserve);
            Elements.Reserves.Clear();
            foreach (String r in strReserves)
            {
                Elements.Add(EEUTILITY.Enums.ClassEnum.Reserve, r);
            }
#endif
            foreach (Reserve r in Elements.Reserves)
            {
                Double dValue = 0.0;
                int mem_id = db.GetMembershipID(EEUTILITY.Enums.CollectionEnum.SystemReserves, Elements.SystemName, r.Name);
                int prop_id = db.PropertyName2EnumId("System", "Reserve", "Reserves", "Type");
                if (db.GetPropertyValue(mem_id, prop_id, 1, ref dValue, null, null, null, null, null, null, null, EEUTILITY.Enums.PeriodEnum.Interval))
                {
                    r.Type = dValue;
                }

                //<-This is ALL commented out because it seems that RemoveProperty and GetPropertiesTable are not working as expected!
                //Remove all Is Enabled properties: 
                //int nRet = 0;
                //int prop_id_enabled = db.PropertyName2EnumId("System", "Reserve", "Reserves", "Is Enabled");

                //(option 1)Remove only untagged data: THIS ONLY WORKS FOR UNTAGGED DATA. IT'S STILL TOO RISKY
                //nRet += db.RemoveProperty(mem_id, prop_id_enabled, 1, null, null, null, null, null, null, null, EEUTILITY.Enums.PeriodEnum.Interval);

                //(option 2)Remove all data: UNFORTUNATELY THIS DOESN'T WORK!!
                //String strPropertyReserve;
                //String strScenarioReserve;
                //ADODB.Recordset rec = db.GetPropertiesTable(EEUTILITY.Enums.CollectionEnum.SystemReserves,null, r,null,null);

                ////Fields in recordset (don't uncomment):
                ////Collection
                ////Parent_x0020_Object
                ////Child_x0020_Object
                ////Property
                ////Value
                ////Data_x0020_File
                ////Units
                ////Band
                ////Date_x0020_From
                ////Date_x0020_To
                ////Timeslice
                ////Action
                ////Expression
                ////Scenario
                ////Memo
                ////Category

                //while (!rec.EOF)
                //{
                //    //foreach (ADODB.Field f in rec.Fields)
                //    //{
                //    //    Console.WriteLine(f.Name);
                //    //}
                //    int nrecord = rec.RecordCount;
                //    strPropertyReserve = rec.Fields["Property"].Value;
                //    if (strPropertyReserve.ToLower() == "is enabled")
                //    {
                //        strScenarioReserve = rec.Fields["Scenario"].Value;
                //        nRet += db.RemoveProperty(mem_id, prop_id_enabled, 1, null, null, null, null, null, strScenarioReserve, null, EEUTILITY.Enums.PeriodEnum.Interval);
                //    }
                //    Console.WriteLine(rec.Fields["Child_x0020_Object"].Value);
                //    rec.MoveNext();
                //}
                //if (nRet > 0)
                //{
                //    Console.WriteLine($"Removed {nRet} properties from {r}");
                //}
                //<-End commented out because RemoveProperty and GetPropertiesTable

            }

            //Membership Generator.Companies:
            String[] GC = db.GetMemberships(EEUTILITY.Enums.CollectionEnum.GeneratorCompanies);
            foreach (String gc in GC)
            {
                String[] par = new string[3] { "Generator (", ").Companies (", ")" };
                String[] s1 = gc.Split(par, StringSplitOptions.RemoveEmptyEntries);

                if (s1.Length == 2)
                {
                    if (strCompanies.Contains(s1[1].Trim()))
                    {
                        Generator g = Elements.GetGenerator(s1[0].Trim());
                        Company c = Elements.GetCompany(s1[1].Trim());
                        c.Generators.Add(g);
                    }
                }
            }
            String[] strScenario = db.GetObjects(EEUTILITY.Enums.ClassEnum.Scenario);
            foreach (String s in strScenario)
            {
                Elements.Add(EEUTILITY.Enums.ClassEnum.Scenario, s);
            }

            //Membership Reserve.Generators:
            String[] RG = db.GetMemberships(EEUTILITY.Enums.CollectionEnum.ReserveGenerators);
            foreach (String rc in RG)
            {
                String[] par = new string[3] { "Reserve (", ").Generators (", ")" };
                String[] s1 = rc.Split(par, StringSplitOptions.RemoveEmptyEntries);

                if (s1.Length == 2)
                {
                    if (Elements.Exists(EEUTILITY.Enums.ClassEnum.Reserve, s1[0].Trim()))
                    {
                        Reserve r = Elements.GetReserve(s1[0].Trim());
                        Generator g = Elements.GetGenerator(s1[1].Trim());
                        r.Generators.Add(g);
                        g.Reserves.Add(r);
                    }
                    //Console.WriteLine($"{r.Name}.{g.Name}");
                }
            }
            //db.Close(); //To save changes
        }

        private static void CreateBackUp()
        {
            strInputFile = CreateBackUpFile(strInputFile); //We switch to the new file in fact
        }

        /// <summary>
        /// Creates a backup file with the same extension but (PIVOT) subfix to name
        /// </summary>
        /// <param name="strFile">Full path to file (including file name and extension)</param>
        public static String CreateBackUpFile(String strFile)
        {
            String strInputFileBackUp = Path.GetDirectoryName(strInputFile) + Path.DirectorySeparatorChar + Path.GetFileNameWithoutExtension(strInputFile) + "(PIVOT)" + Path.GetExtension(strInputFile);
            File.Copy(strInputFile, strInputFileBackUp, true);
            Console.WriteLine("Backing up file " + strInputFile);
            return strInputFileBackUp;
        }

        private static bool CheckExistance()
        {
            List<String> lDBReserves = new List<String>(db.GetObjects(EEUTILITY.Enums.ClassEnum.Reserve));
            foreach (Reserve r in Elements.Reserves)
            {
                if (!lDBReserves.Contains(r.Name))
                {
                    throw new Exception("Reserve '" + r.Name + "' defined in config file must exist in database " + db.InstallPath);
                }
            }
            List<String> lDBCompanies = new List<String>(db.GetObjects(EEUTILITY.Enums.ClassEnum.Company));
            foreach (Company c in Elements.Companies)
            {
                if (!lDBCompanies.Contains(c.Name))
                {
                    throw new Exception("Company '" + c.Name + "' defined in config file must exist in database " + db.InstallPath);
                }
            }
            List<String> lDBHorizons = new List<String>(db.GetObjects(EEUTILITY.Enums.ClassEnum.Horizon));
            foreach (String strHorizon in lconfig_base_horizons)
            {
                if (!lDBHorizons.Contains(strHorizon))
                {
                    throw new Exception("Horizon '" + strHorizon + "' defined in config file must exist in database " + db.InstallPath);
                }
            }
            List<String> lDBScenarios = new List<String>(db.GetObjects(EEUTILITY.Enums.ClassEnum.Scenario));
            foreach (String strScenario in ModelBase.Scenarios)
            {
                if (!lDBScenarios.Contains(strScenario))
                {
                    throw new Exception("Scenario '" + strScenario + "' defined in config file must exist in database " + db.InstallPath);
                }
            }
            //List<String> lDBModels = new List<String>(db.GetObjects(EEUTILITY.Enums.ClassEnum.Model));
            //if (!lDBModels.Contains(ModelBase.Name))
            //{
            //    throw new Exception($"Model base {ModelBase.Name} not found in this database! Adjust config file or change database");
            //}
            //TODO: AGREGAR UNA VALIDACIÓN DE EXISTENCIA DE LOS MODELOS?
            return true;
        }

        /// <summary>
        /// Checks if the file in the argument checks some minimum requirements
        /// </summary>
        /// <param name="strFile">Full path to file (including file name and extension)</param>
        /// <returns></returns>
        public static bool isValidDB(String strFile)
        {
            if (!File.Exists(strFile))
            {
                throw new Exception($"PLEXOS Input xml file {strFile} doesn't exist or is not located in the given path");
            }
            if (!Path.GetExtension(strFile).ToLower().Equals(".xml"))
            {
                throw new Exception($"Invalid PLEXOS input file type. Only xml files are supported");
            }
            String strConfigFile = Path.GetDirectoryName(strFile) + Path.DirectorySeparatorChar + CONFIG_FILE;
            if (!File.Exists(strConfigFile))
            {
                throw new Exception($"Configuration file couldn't be found! Configuration file should be in the same directory as PLEXOS input file and named 'config.json'");
            }
            return true;
        }

        private static void PutAssemblies()
        {
            bool hasAssembly = false;
            foreach (String strAssembly in db.GetAssemblies())
            {
                if (hasAssembly = (strAssembly == ASSEMBLY_FILE))
                {
                    break;
                }
            }
            if (!hasAssembly)
            {
                db.AddAssembly(ASSEMBLY_FILE, ASSEMBLY_NAMESPACE);
                db.Close();
            }
        }

        private static void PutGeneratorProperties()
        {
            //Delete Unit Commitment Optimality values (untagged data) This function doesn't work!!
            int nRet = 0;
            int prop_id = db.PropertyName2EnumId("System", "Generator", "Generators", "Unit Commitment Optimality");
            foreach (Generator g in Elements.Generators)
            {
                int mem_id = db.GetMembershipID(EEUTILITY.Enums.CollectionEnum.SystemGenerators, Elements.SystemName, g.Name);
                nRet += db.RemoveProperty(mem_id, prop_id, 1, null, null, null, null, null, null, null, EEUTILITY.Enums.PeriodEnum.Interval); //This is not working!!!
                //Double dVal = -999999;
                //if (db.GetPropertyValue(mem_id, prop_id, 1, ref dVal, null, null, null, null, null, null, null, 0))
                //{
                //    Console.WriteLine($"Total {nRet} 'Unit Commitment Optimality' property entries in database"); //Not working either!
                //}
            }
            db.Close();
            if (nRet > 0)
            {
                Console.WriteLine($"Deleted {nRet} 'Unit Commitment Optimality' property entries from database");
            }
        }

        private static void PutReserveProperties()
        {
            //DEPRECATED: RemoveProperty is not working properly
            //int nAdd = 0;
            //int nRet = 0;
            //foreach (Reserve reserve in Elements.Reserves)
            //{
            //    String strScenario = reserve.Name + "=0";
            //    Debug.Assert(Elements.Exists(EEUTILITY.Enums.ClassEnum.Scenario, strScenario), "You need to create all scenarios first!");
            //    int mem_id = db.GetMembershipID(EEUTILITY.Enums.CollectionEnum.SystemReserves, Elements.SystemName, reserve.Name);
            //    int prop_id = db.PropertyName2EnumId("System", "Reserve", "Reserves", "Min Provision");
            //    nAdd += db.AddProperty(mem_id, prop_id, 1, 0.0, null, null, null, null, null, strScenario, null, EEUTILITY.Enums.PeriodEnum.Interval);
            //    prop_id = db.PropertyName2EnumId("System", "Reserve", "Reserves", "Is Enabled");
            //    nRet = db.RemoveProperty(mem_id, prop_id, 1, null, null, null, null, null, null, null, EEUTILITY.Enums.PeriodEnum.Interval);
            //}
            //if (nRet > 0)
            //{
            //    Console.WriteLine($"Deleted {nRet} 'Is Enabled' property entries from Reserve objects in database");
            //}
            //foreach (Reserve reserve in Elements.Reserves)
            //{
            //    String strScenario = reserve.Name + "_ON";
            //    Debug.Assert(Elements.Exists(EEUTILITY.Enums.ClassEnum.Scenario, strScenario), "You need to create all scenarios first!");
            //    int mem_id = db.GetMembershipID(EEUTILITY.Enums.CollectionEnum.SystemReserves, Elements.SystemName, reserve.Name);
            //    int prop_id = db.PropertyName2EnumId("System", "Reserve", "Reserves", "Is Enabled");
            //    nAdd += db.AddProperty(mem_id, prop_id, 1, -1, null, null, null, null, null, strScenario, null, EEUTILITY.Enums.PeriodEnum.Interval); //true
            //}
            //db.Close();
        }

        private static void PutReserveGeneratorProperties()
        {
            foreach (Company company in Elements.Companies)
            {
                Console.WriteLine($"  Adding properties for pivotal company {company.Name}");
                foreach (Reserve reserve in Elements.Reserves)
                {
                    PivotalReserveCapacity(company, reserve, "Max Response");
                    if (reserve.Type == 6) //Operational
                    {
                        PivotalReserveCapacity(company, reserve, "Max Replacement");
                    }
                }
            }
            db.Close();
        }

        private static void PivotalReserveCapacity(Company company, Reserve reserve, String strProperty)
        {
            String strScenario = MakeScenarioLabel(company, reserve);
            Debug.Assert(Elements.Exists(EEUTILITY.Enums.ClassEnum.Scenario, strScenario), "You need to create all scenarios first!");
            try
            {
                foreach (Generator g in company.Generators)
                {
                    if (g.Reserves.Contains(reserve))
                    {
                        int mem_id = db.GetMembershipID(EEUTILITY.Enums.CollectionEnum.ReserveGenerators, reserve.Name, g.Name);
                        int prop_id = db.PropertyName2EnumId("Reserve", "Generator", "Generators", strProperty);
                        int ret = db.AddProperty(mem_id, prop_id, 1, MAX_RESPONSE_ZERO, null, null, null, null, null, strScenario, null, EEUTILITY.Enums.PeriodEnum.Interval);
                        //Console.WriteLine($"{company.Name}.{g.Name}");
                    }
                }
                //db.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("Error adding properties: " + e.Message);
            }
        }

        private static void PutDecisionVariableProperties()
        {
            int nAttributeType = GetAttributeEnum(db, "Decision Variable", "Type");
            foreach (String dv in db.GetObjects(EEUTILITY.Enums.ClassEnum.DecisionVariable))
            {
                db.SetAttributeValue(EEUTILITY.Enums.ClassEnum.DecisionVariable, dv, nAttributeType, 0); //All continuos!
            }
        }

        private static void PutModelProperties()
        {
            int nAttributeEnabled = GetAttributeEnum(db, "Model", "Enabled");
            int nAttributeAssembly = GetAttributeEnum(db, "Model", "Load Custom Assemblies");
            foreach (String strModelName in db.GetObjects(EEUTILITY.Enums.ClassEnum.Model))
            {
                db.SetAttributeValue(EEUTILITY.Enums.ClassEnum.Model, strModelName, nAttributeEnabled, 0); //this is a false
            }
            foreach (String strModelNameNew in lconfig_models_new)
            {
                db.SetAttributeValue(EEUTILITY.Enums.ClassEnum.Model, strModelNameNew, nAttributeEnabled, -1); //this is a true
                db.SetAttributeValue(EEUTILITY.Enums.ClassEnum.Model, strModelNameNew, nAttributeAssembly, -1); //this is a true
            }
            db.Close();
        }

        /// <summary>
        /// Don't use! Unfortunately, configuring report is not possible with current Plexos version. AddReportProperty can't be used to configure report properties
        /// </summary>
        private static void PutReportProperties()
        {
            int nAdd = 0;
            foreach (ReportProperty r in lconfig_base_reports)
            {
                //int nReportPropertyId = db.ReportPropertyName2PropertyId(r.parent_class, r.child_class, r.collection, r.property);
                int nReportPropertyId = db.ReportPropertyName2EnumId(r.parent_class, r.child_class, r.collection, r.property);
                try
                {
                    db.AddReportProperty(ModelBase.Report, nReportPropertyId, EEUTILITY.Enums.SimulationPhaseEnum.MTSchedule, true, false, true, false);
                    nAdd++;
                    Console.WriteLine($"Added '{r.property}' report properties to database"); //for debugging only
                } catch (Exception e)
                {
                    continue;
                }
            }
            if (nAdd > 0)
            {
                Console.WriteLine($"Added {nAdd} report properties to database");
                db.Close();
            }
        }

        private static void PutScenarioProperties()
        {
            //We only need to push back the reading order of other scenarios and proritize the pivot scenarios
            int nAttributeOrder = GetAttributeEnum(db, "Scenario", "Read Order");
            foreach (String strBaseScenarios in ModelBase.Scenarios)
            {
                db.SetAttributeValue(EEUTILITY.Enums.ClassEnum.Scenario, strBaseScenarios, nAttributeOrder, 0); //We lower the priority of base scenarios
            }
        }

        private static void PutDiagnosticProperties()
        {
            int nAttributeEnabled;
            nAttributeEnabled = GetAttributeEnum(db, "Diagnostic", "LP Files");
            db.SetAttributeValue(EEUTILITY.Enums.ClassEnum.Diagnostic, ModelBase.Diagnostic.Name, nAttributeEnabled, ModelBase.Diagnostic.LPFiles);
            nAttributeEnabled = GetAttributeEnum(db, "Diagnostic", "Objective Function");
            db.SetAttributeValue(EEUTILITY.Enums.ClassEnum.Diagnostic, ModelBase.Diagnostic.Name, nAttributeEnabled, ModelBase.Diagnostic.ObjectiveFunction);
            nAttributeEnabled = GetAttributeEnum(db, "Diagnostic", "LP Progress");
            db.SetAttributeValue(EEUTILITY.Enums.ClassEnum.Diagnostic, ModelBase.Diagnostic.Name, nAttributeEnabled, ModelBase.Diagnostic.LPProgress);
            nAttributeEnabled = GetAttributeEnum(db, "Diagnostic", "MIP Progress");
            db.SetAttributeValue(EEUTILITY.Enums.ClassEnum.Diagnostic, ModelBase.Diagnostic.Name, nAttributeEnabled, ModelBase.Diagnostic.MIPProgress);
            nAttributeEnabled = GetAttributeEnum(db, "Diagnostic", "Times");
            db.SetAttributeValue(EEUTILITY.Enums.ClassEnum.Diagnostic, ModelBase.Diagnostic.Name, nAttributeEnabled, ModelBase.Diagnostic.Times);
            nAttributeEnabled = GetAttributeEnum(db, "Diagnostic", "Step Summary");
            db.SetAttributeValue(EEUTILITY.Enums.ClassEnum.Diagnostic, ModelBase.Diagnostic.Name, nAttributeEnabled, ModelBase.Diagnostic.StepSummary);
            nAttributeEnabled = GetAttributeEnum(db, "Diagnostic", "Performance Summary");
            db.SetAttributeValue(EEUTILITY.Enums.ClassEnum.Diagnostic, ModelBase.Diagnostic.Name, nAttributeEnabled, ModelBase.Diagnostic.PerformanceSummary);
            nAttributeEnabled = GetAttributeEnum(db, "Diagnostic", "Task Size");
            db.SetAttributeValue(EEUTILITY.Enums.ClassEnum.Diagnostic, ModelBase.Diagnostic.Name, nAttributeEnabled, ModelBase.Diagnostic.TaskSize);
            nAttributeEnabled = GetAttributeEnum(db, "Diagnostic", "Database Load");
            db.SetAttributeValue(EEUTILITY.Enums.ClassEnum.Diagnostic, ModelBase.Diagnostic.Name, nAttributeEnabled, ModelBase.Diagnostic.DatabaseLoad);
            nAttributeEnabled = GetAttributeEnum(db, "Diagnostic", "Data File Read");
            db.SetAttributeValue(EEUTILITY.Enums.ClassEnum.Diagnostic, ModelBase.Diagnostic.Name, nAttributeEnabled, ModelBase.Diagnostic.DataFileRead);
            nAttributeEnabled = GetAttributeEnum(db, "Diagnostic", "Computer Information");
            db.SetAttributeValue(EEUTILITY.Enums.ClassEnum.Diagnostic, ModelBase.Diagnostic.Name, nAttributeEnabled, ModelBase.Diagnostic.ComputerInformation);
            db.Close();
        }

        private static void PutProperties()
        {
            Console.WriteLine("Putting Generator properties..");
            PutGeneratorProperties();

            Console.WriteLine($"Putting Reserve properties..");
            PutReserveProperties();

            Console.WriteLine($"Putting Reserve-Generators properties..");
            PutReserveGeneratorProperties();

            Console.WriteLine($"Putting Decision Variable properties..");
            PutDecisionVariableProperties();

            Console.WriteLine($"Putting Model properties..");
            PutModelProperties();

            //Console.WriteLine($"Putting Report properties.."); //Broken!
            //PutReportProperties();

            Console.WriteLine($"Putting Scenario properties..");
            PutScenarioProperties();

            Console.WriteLine($"Putting Diagnostic properties..");
            PutDiagnosticProperties();
        }

        /// <summary>
        /// DEPRECATED: THis method is not in use because reserves use their own existent scenario now
        /// </summary>
        private static void CreateReserveScenarios()
        {
            if (!db.CategoryExists(EEUTILITY.Enums.ClassEnum.Scenario, RESERVE_CATEGORY))
            {
                db.AddCategory(EEUTILITY.Enums.ClassEnum.Scenario, RESERVE_CATEGORY);
            }
            foreach (Reserve reserve in Elements.Reserves)
            {
                String strScenario = reserve.Name + "_ON";
                if (!Elements.Exists(EEUTILITY.Enums.ClassEnum.Scenario, strScenario))
                {
                    int nScenario = db.AddObject(strScenario, EEUTILITY.Enums.ClassEnum.Scenario, true, RESERVE_CATEGORY, $"Scenenario for properties where reserve{reserve.Name} is active (or not)");
                    Elements.Add(EEUTILITY.Enums.ClassEnum.Scenario, strScenario);
                }
            }
            db.Close();
        }

        private static void CreatePivotScenarios()
        {
            int nAttributeOrder = GetAttributeEnum(db, "Scenario", "Read Order");
            foreach (Company company in Elements.Companies)
            {
                if (!db.CategoryExists(EEUTILITY.Enums.ClassEnum.Scenario, company.Name))
                {
                    db.AddCategory(EEUTILITY.Enums.ClassEnum.Scenario, company.Name);
                }
                foreach (Reserve reserve in Elements.Reserves)
                {
                    String strScenario = MakeScenarioLabel(company, reserve);
                    if (!Elements.Exists(EEUTILITY.Enums.ClassEnum.Scenario, strScenario))
                    {
                        int nScenario = db.AddObject(strScenario, EEUTILITY.Enums.ClassEnum.Scenario, true, company.Name, $"Properties for pivotal company '{company.Name}' for reserve service '{reserve.Name}'");
                        Elements.Add(EEUTILITY.Enums.ClassEnum.Scenario, strScenario);
                        db.SetAttributeValue(EEUTILITY.Enums.ClassEnum.Scenario, strScenario, nAttributeOrder, 2); //Raise priority of this scenario
                    }
                }
            }
            db.Close();
        }

        private static String MakeScenarioLabel(Company company, Reserve reserve)
        {
            return "Pivot" + company.Name + "_" + reserve.Name;
        }

        private static void CreateScenarios()
        {
            //CreateReserveScenarios();
            CreatePivotScenarios();
        }

        private static void CreateBaseModels()
        {
            if (!db.CategoryExists(EEUTILITY.Enums.ClassEnum.Model, ModelBase.Name))
            {
                db.AddCategory(EEUTILITY.Enums.ClassEnum.Model, ModelBase.Name);
            }
            foreach (String strHorizon in lconfig_base_horizons)
            {
                //Create Object and AddMemberships
                String strModelName = ModelBase.Name + strHorizon;
                int nObj = db.AddObject(strModelName, EEUTILITY.Enums.ClassEnum.Model, true, ModelBase.Name);
                if (nObj > 0)
                {
                    lconfig_models_new.Add(strModelName);
                }
                db.AddMembership(EEUTILITY.Enums.CollectionEnum.ModelHorizon, strModelName, strHorizon);
                AddMembershipsModelBase(db, strModelName);

                //Add reserve activation scenarios:
                foreach (Reserve r in Elements.Reserves)
                {
                    ////String strReserveOnScenario = r.Name + "_ON";
                    //db.AddMembership(EEUTILITY.Enums.CollectionEnum.ModelScenarios, strModelName, strReserveOnScenario);
                    foreach (String s in r.Scenarios)
                    {
                        db.AddMembership(EEUTILITY.Enums.CollectionEnum.ModelScenarios, strModelName, s);
                    }
                }

                //Put the attribute Load Custom Assemblies
                int nAttributeLoad = GetAttributeEnum(db, "Model", "Load Custom Assemblies");
                db.SetAttributeValue(EEUTILITY.Enums.ClassEnum.Model, strModelName, nAttributeLoad, -1); //this is a true
            }
            db.Close();
        }

        private static void CreateReserve0Models()
        {
            String strCategoryName = ModelBase.Name + "_R=0";
            if (!db.CategoryExists(EEUTILITY.Enums.ClassEnum.Model, strCategoryName))
            {
                db.AddCategory(EEUTILITY.Enums.ClassEnum.Model, strCategoryName);
            }
            foreach (String strHorizon in lconfig_base_horizons)
            {
                foreach (Reserve reserve in Elements.Reserves)
                {
                    //String strModelBase = ModelBase.Name + strHorizon;
                    String strModelNew0 = ModelBase.Name + strHorizon + "_" + reserve.Name + "=0";

                    //db.CopyObject(strModelBase, strModelNew0, EEUTILITY.Enums.ClassEnum.Model
                    //Workaround to CopyObject:
                    int nObj = db.AddObject(strModelNew0, EEUTILITY.Enums.ClassEnum.Model, true, strCategoryName, $"Model where reserve{reserve.Name}.[Min Provision] = 0 and Horizon {strHorizon}");
                    if (nObj > 0)
                    {
                        lconfig_models_new.Add(strModelNew0);
                    }
                    db.AddMembership(EEUTILITY.Enums.CollectionEnum.ModelHorizon, strModelNew0, strHorizon);
                    AddMembershipsModelBase(db, strModelNew0);
                    //End workaround

                    //Add reserve activation scenarios (all of them except the one we want to remove):
                    //This is actually a workaround to another issue with the API
                    foreach (Reserve r in Elements.Reserves)
                    {
                        if (r.Name != reserve.Name)
                        {
                            //String strReserveOnScenario = r.Name + "_ON";
                            foreach (String s in r.Scenarios)
                            {
                                db.AddMembership(EEUTILITY.Enums.CollectionEnum.ModelScenarios, strModelNew0, s);
                            }
                        }
                    }

                    db.CategorizeObject(EEUTILITY.Enums.ClassEnum.Model, strModelNew0, strCategoryName);

                    //Put the attribute Load Custom Assemblies
                    int nAttributeLoad = GetAttributeEnum(db, "Model", "Load Custom Assemblies");
                    db.SetAttributeValue(EEUTILITY.Enums.ClassEnum.Model, strModelNew0, nAttributeLoad, 0); //this is a false
                }
            }
            db.Close();
        }

        private static void CreatePivotModels()
        {
            //Pivotal models:
            foreach (Company company in Elements.Companies)
            {
                String strCategory = company.Name;
                if (!db.CategoryExists(EEUTILITY.Enums.ClassEnum.Model, strCategory))
                {
                    db.AddCategory(EEUTILITY.Enums.ClassEnum.Model, strCategory);
                }
                
                foreach (Reserve reserve in Elements.Reserves)
                {
                    String strScenarioPivot = MakeScenarioLabel(company, reserve);
                    foreach (String strHorizon in lconfig_base_horizons)
                    {
                        //String strModelBase = ModelBase.Name + strHorizon;
                        String strModelPivot = ModelBase.Name + strHorizon + "_" + company.Name + "_" + reserve.Name;

                        //db.CopyObject(strModelBase, strModelPivot, EEUTILITY.Enums.ClassEnum.Model);
                        //db.CategorizeObject(EEUTILITY.Enums.ClassEnum.Model, strModelPivot, company.Name);

                        //Workaround to CopyObject:
                        int nObj = db.AddObject(strModelPivot, EEUTILITY.Enums.ClassEnum.Model, true, strCategory, $"Model for pivotal company {company.Name}: Reserve '{reserve.Name}' and Horizon {strHorizon}");
                        if (nObj > 0)
                        {
                            lconfig_models_new.Add(strModelPivot);
                        }
                        db.AddMembership(EEUTILITY.Enums.CollectionEnum.ModelHorizon, strModelPivot, strHorizon);
                        AddMembershipsModelBase(db, strModelPivot);
                        //End workaround
                        
                        db.AddMembership(EEUTILITY.Enums.CollectionEnum.ModelScenarios, strModelPivot, strScenarioPivot);

                        //Add reserve activation scenarios (all of them):
                        foreach (Reserve r in Elements.Reserves)
                        {
                            //String strReserveOnScenario = r.Name + "_ON";
                            //db.AddMembership(EEUTILITY.Enums.CollectionEnum.ModelScenarios, strModelPivot, strReserveOnScenario);
                            foreach (String s in r.Scenarios)
                            {
                                db.AddMembership(EEUTILITY.Enums.CollectionEnum.ModelScenarios, strModelPivot, s);
                            }
                        }

                        //int nAttribute = db.PropertyName2EnumId("System", "Model", "Models", "Enabled"); //Not working! It's an attribute!
                        //int nAttributeEnabled = GetAttributeEnum(db, "Model", "Enabled");
                        //db.SetAttributeValue(EEUTILITY.Enums.ClassEnum.Model, strModelPivot, nAttributeEnabled, -1); //this is a true (moved to PutProperties)
                        int nAttributeLoad = GetAttributeEnum(db, "Model", "Load Custom Assemblies");
                        db.SetAttributeValue(EEUTILITY.Enums.ClassEnum.Model, strModelPivot, nAttributeLoad, 0); //this is a false

                    }
                }
            }
            db.Close();
        }

        private static void CreateModels()
        {
            CreateBaseModels();
            CreateReserve0Models();
            CreatePivotModels();
        }

        /// <summary>
        /// Don't use! Unfortunately, configuring report is not possible with current Plexos version. AddReportProperty can't be used to configure report properties
        /// </summary>
        private static void CreateReports()
        {
            String[] reports = db.GetObjects(EEUTILITY.Enums.ClassEnum.Report);
            foreach (String r in reports)
            {
                if (r == ModelBase.Report)
                {
                    return;
                }
            }
            db.AddObject(ModelBase.Report, EEUTILITY.Enums.ClassEnum.Report, true);
        }

        private static void CreateDiagnostics()
        {
            String[] diagnostics = db.GetObjects(EEUTILITY.Enums.ClassEnum.Diagnostic);
            foreach (String d in diagnostics)
            {
                if (d == ModelBase.Diagnostic.Name)
                {
                    return;
                }
            }
            db.AddObject(ModelBase.Diagnostic.Name, EEUTILITY.Enums.ClassEnum.Diagnostic, true);
        }

        private static void AddMembershipsModelBase(PLEXOS7_NET.Core.DatabaseCore db, String strModelName)
        {
            db.AddMembership(EEUTILITY.Enums.CollectionEnum.ModelReport, strModelName, ModelBase.Report);
            db.AddMembership(EEUTILITY.Enums.CollectionEnum.ModelMTSchedule, strModelName, ModelBase.MT);
            db.AddMembership(EEUTILITY.Enums.CollectionEnum.ModelTransmission, strModelName, ModelBase.Transmission);
            db.AddMembership(EEUTILITY.Enums.CollectionEnum.ModelProduction, strModelName, ModelBase.Production);
            //db.AddMembership(EEUTILITY.Enums.CollectionEnum.ModelStochastic, strModelName, ModelBase.Stochastic); //Deprecated for the moment
            db.AddMembership(EEUTILITY.Enums.CollectionEnum.ModelPerformance, strModelName, ModelBase.Performance);
            db.AddMembership(EEUTILITY.Enums.CollectionEnum.ModelDiagnostic, strModelName, ModelBase.Diagnostic.Name);
            foreach (String strScenario in ModelBase.Scenarios)
            {
                db.AddMembership(EEUTILITY.Enums.CollectionEnum.ModelScenarios, strModelName, strScenario);
            }
        }

        private static int GetAttributeEnum(PLEXOS7_NET.Core.DatabaseCore db, String strClassName, String strAttributeName)
        {
            //String[] sClassFields = new String[] { "class_id", "name", "enum_id" };
            String[] sClassFields = new String[] { }; //Workaround! this seems to work the opposite! it returns all fields except the selected!
            int nClassId = GetClassId(db, strClassName);
            ADODB.Recordset rec = db.GetData("t_attribute", ref sClassFields);

            while (!rec.EOF)
            {
                int nId = rec.Fields["class_id"].Value;
                String strName = rec.Fields["name"].Value;
                if (nId == nClassId && strName == strAttributeName)
                {
                    int nRet = rec.Fields["enum_id"].Value;
                    rec.Close();
                    return nRet;
                }
                rec.MoveNext();
            }
            rec.Close();
            throw new Exception("Couldn't find the id for class " + strClassName);
        }

        private static Dictionary<String, int> ClassIds = new Dictionary<string, int>();
        private static int GetClassId(PLEXOS7_NET.Core.DatabaseCore db, String strClassName)
        {
            if (ClassIds.Count == 0)
            {
                InitClassId(db);
            }
            return ClassIds[strClassName];
        }

        private static void InitClassId(PLEXOS7_NET.Core.DatabaseCore db)
        {
            //String[] sClassFields = new String[] {"class_id", "name"};
            String[] sClassFields = new String[] { }; //Workaround! this seems to work the opposite! it returns all fields except the selected!
            ADODB.Recordset rec = db.GetData("t_class", ref sClassFields);
            while (!rec.EOF)
            {
                String strName = rec.Fields["name"].Value;
                int nRet = rec.Fields["class_id"].Value;
                ClassIds.Add(strName, nRet);
                rec.MoveNext();
            }
            rec.Close();
        }

        #endregion

        #region Solution Calculations

        private static void InitSolution()
        {
            zip = new PLEXOS7_NET.Core.Solution();
            zip.Connection(strOutputFile);
            Elements.SystemName = zip.SystemName;

            String[] strGens = zip.GetObjects(EEUTILITY.Enums.ClassEnum.Generator);
            foreach (String g in strGens)
            {
                Elements.Add(EEUTILITY.Enums.ClassEnum.Generator, g);
            }
            String[] strRes = zip.GetObjects(EEUTILITY.Enums.ClassEnum.Reserve);
            foreach (String r in strRes)
            {
                Elements.Add(EEUTILITY.Enums.ClassEnum.Reserve, r);
            }

            //nothing else for the moment

        }

        /// <summary>
        /// Simple existence check and zip file extension
        /// </summary>
        /// <param name="strFile"></param>
        /// <returns></returns>
        public static bool isValidZip(String strFile)
        {
            if (!File.Exists(strFile))
            {
                throw new Exception($"PLEXOS output zip file {strFile} doesn't exist or is not located in the given path");
            }
            if (!Path.GetExtension(strFile).ToLower().Equals(".zip"))
            {
                throw new Exception($"Invalid PLEXOS output zip file type. Only zip files are supported");
            }
            return true;
        }

        private static EEUTILITY.Enums.CollectionEnum FixCollectionId(String parentClass, String childClass, String collection)
        {
            int id = zip.CollectionName2Id(parentClass, childClass, collection);
            foreach (EEUTILITY.Enums.CollectionEnum e in Enum.GetValues(typeof(EEUTILITY.Enums.CollectionEnum)).Cast<EEUTILITY.Enums.CollectionEnum>())
            {
                int nEnum = (int)e;
                if (nEnum == id)
                {
                    return e;
                }
            }
            throw new Exception($"Can't find enum for {parentClass}.{childClass}.{collection}!");
        }

        private static String FixPropertyId(String parentClass, String childClass, String collection, String property)
        {
            int id = zip.PropertyName2EnumId(parentClass, childClass, collection, property);
            foreach (EEUTILITY.Enums.SystemOutGeneratorsEnum e in Enum.GetValues(typeof(EEUTILITY.Enums.SystemOutGeneratorsEnum)).Cast<EEUTILITY.Enums.SystemOutGeneratorsEnum>())
            {
                int nEnum = (int)e;
                if (nEnum == id)
                {
                    return id.ToString();
                }
            }
            throw new Exception($"Can't find enum for {parentClass}.{childClass}.{collection}.{property}!");
        }

        private static double QryPropertyGenerators(String property)
        {
            ADODB.Recordset rec;
            double dTotalValue = 0.0;
            String strPropertyList = FixPropertyId("System", "Generator", "Generators", property);
            EEUTILITY.Enums.CollectionEnum collection = FixCollectionId("System", "Generator", "Generators");

            rec = zip.Query(EEUTILITY.Enums.SimulationPhaseEnum.MTSchedule, collection, zip.SystemName, null, EEUTILITY.Enums.PeriodEnum.Interval, EEUTILITY.Enums.SeriesTypeEnum.Properties, strPropertyList);
            int cont = 0;
            while (!rec.EOF)
            {
                double dValue = Double.Parse(rec.Fields["value"].Value.ToString());
                dTotalValue += dValue;
                //Console.WriteLine($"{g.Name}: {dValue}");
                cont++;
                rec.MoveNext();
            }
            Console.WriteLine($" TOTAL: {dTotalValue}");
            rec.Close();
            return dTotalValue;
        }

        #endregion

        #region Config File Operations

        private static void InitConfig()
        {
            String strConfigFile =Path.GetDirectoryName(strInputFile) + Path.DirectorySeparatorChar + CONFIG_FILE;
            
            dynamic text = JsonConvert.DeserializeObject(File.ReadAllText(strConfigFile));
            //jlist2List(text.Reserve, lconfig_reserves); //Reserves:
            jlist2List(text.Company, lconfig_companies); //Companies:
            jlist2List(text.Horizon, lconfig_base_horizons); //Horizons:
            jlist2List(text.Scenario, lconfig_base_scenarios); //Scenarios:
            //lconfig_base_reports = JsonConvert.DeserializeObject<List<ReportProperty>>(text.Report); //Not sure what i'm doing wrong here!

            //Workaround to reserves:
            for (int i = 0; i < text.Reserve.Count; i++)
            {
                String strReserveName = text.Reserve[i].Name;
                Reserve r = Elements.AddReserve(strReserveName);
                if (text.Reserve[i].Scenario != null)
                {
                    string s = text.Reserve[i].Scenario;
                    r.Scenarios.Add(s);
                }
                if (text.Reserve[i].Scenarios != null)
                {
                    for (int k = 0; k < text.Reserve[i].Scenarios.Count; k++)
                    {
                        string s = text.Reserve[i].Scenarios[k];
                        r.Scenarios.Add(s);
                    }
                }
            }

            //Workaround to reports:
            for (int i = 0; i < text.Report.Count; i++)
            {
                ReportProperty r = new ReportProperty();
                r.parent_class = text.Report[i].parent_class;
                r.child_class = text.Report[i].child_class;
                r.collection = text.Report[i].collection;
                r.property = text.Report[i].property;
                lconfig_base_reports.Add(r);
            }
            //End workaround

            //Model:
            ModelBase = new Model
            {
                Name = text.Model.Name,
                Horizon = text.Model.Horizon,
                Report = text.Model.Report,
                MT = text.Model.MT,
                Transmission = text.Model.Transmission,
                Production = text.Model.Production,
                Stochastic = text.Model.Stochastic,
                Performance = text.Model.Performance,
            };
            //Base Scenarios:
            foreach (String s in lconfig_base_scenarios)
            {
                ModelBase.Scenarios.Add(s);
            }

            //Diagnostic:
            Diagnostic d = new Diagnostic
            {
                Name = text.Diagnostic.Name,
                LPFiles = text.Diagnostic.LPFiles,
                ObjectiveFunction = text.Diagnostic.ObjectiveFunction,
                LPProgress = text.Diagnostic.LPProgress,
                MIPProgress = text.Diagnostic.MIPProgress,
                Times = text.Diagnostic.Times,
                StepSummary = text.Diagnostic.StepSummary,
                PerformanceSummary = text.Diagnostic.PerformanceSummary,
                TaskSize = text.Diagnostic.TaskSize,
                DatabaseLoad = text.Diagnostic.DatabaseLoad,
                DataFileRead = text.Diagnostic.DataFileRead,
                ComputerInformation = text.Diagnostic.ComputerInformation
            };
            ModelBase.Diagnostic= d;

        }

        static private void jlist2List(dynamic d, List<String> l)
        {
            int n = d.Count;
            for (int i = 0; i < n; i++)
            {
                string s = d[i];
                l.Add(s);
            }
        }

        #endregion
        
        [STAThread]
        static void Main(string[] args)
        {

            strInputFile = "";
            //System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance); //Necessary for net core

            //Check if there's a valid argument file:
            if (args.Length == 1)
            {
                if (!File.Exists(args[0]))
                {
                    Console.WriteLine("Can't find file " + args[0]);
                    Console.ReadKey();
                    System.Environment.Exit(1);
                } else
                {
                    strInputFile = args[0];
                }
            }

            //If the argument didn't work, we open a file dialog to ask for the file: (only .net!)
            if (strInputFile == "")
            {
                OpenFileDialog ofd = new OpenFileDialog();
                DialogResult dr = ofd.ShowDialog();
                if (dr == DialogResult.OK)
                {
                    strInputFile = ofd.FileName;
                }
                else
                {
                    Console.WriteLine("Execution Cancelled. No input file selected");
                    Console.WriteLine($"Press any key to finish...");
                    Console.ReadKey();
                    System.Environment.Exit(1);
                }
            }

            //Now we can now execute everything:
            if (isValidDB(strInputFile))
            {
                InitConfig();
                CreateBackUp();
                InitConnection();
                CheckExistance();
                InitInput();
                CreateScenarios();
                CreateDiagnostics();
                //CreateReports(); //Broken!
                CreateModels();
                PutProperties();
                PutAssemblies();

                Console.WriteLine("Finished updating input file. Press any key to finish...");
                Console.ReadKey();
            } else if (isValidZip(strInputFile))
            {
                //OUTPUT RUTINES: DEPRECATED! Version 8.1 has a major bug with the reference's version. This entire project was moved to Excel COM library
                strOutputFile = strInputFile;
                InitSolution();
                QryPropertyGenerators("Generation");
                QryPropertyGenerators("Generation Cost");
                Console.WriteLine("Finished reading output file. Press any key to finish...");
                Console.ReadKey();
                throw new NotImplementedException("Function not yet implemented in .NET");
            }
            else
            {
                Console.WriteLine($"Invalid file path or unsupported input file {strInputFile}");
                Console.WriteLine($"Press any key to finish...");
                Console.ReadKey();
            }
        }
    }

}
