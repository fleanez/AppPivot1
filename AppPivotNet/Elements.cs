using System;
using System.Collections.Generic;

namespace AppPivotNet
{
    class Elements
    {
        private static Dictionary<String, Company> _Companies = new Dictionary<string, Company>();
        private static Dictionary<String, Generator> _Generators = new Dictionary<string, Generator>();
        private static Dictionary<String, Reserve> _Reserves = new Dictionary<string, Reserve>();
        private static Dictionary<String, Scenario> _Scenarios = new Dictionary<string, Scenario>();

        public static void Add(EEUTILITY.Enums.ClassEnum classEnum, String name)
        {
            switch (classEnum)
            {
                case EEUTILITY.Enums.ClassEnum.Company:
                    AddCompany(name);
                    break;
                case EEUTILITY.Enums.ClassEnum.Generator:
                    AddGenerator(name);
                    break;
                case EEUTILITY.Enums.ClassEnum.Reserve:
                    AddReserve(name);
                    break;
                case EEUTILITY.Enums.ClassEnum.Scenario:
                    AddScenario(name);
                    break;
                default:
                    break;
            }
        }

        /// <summary>
        /// Creates and registers a new company Element
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static Company AddCompany(String name)
        {
            Company c = new Company();
            c.Name = name;
            _Companies.Add(name, c);
            return c;
        }
        /// <summary>
        /// Creates and registers a new Generator Element
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static Generator AddGenerator(String name)
        {
            Generator g = new Generator(name);
            _Generators.Add(name, g);
            return g;
        }
        /// <summary>
        /// Creates and registers a new Reserve Element
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static Reserve AddReserve(String name)
        {
            Reserve r = new Reserve();
            r.Name = name;
            _Reserves.Add(name, r);
            return r;
        }
        /// <summary>
        /// Creates and registers a new Scenario Element
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static Scenario AddScenario(String name)
        {
            Scenario s = new Scenario();
            s.Name = name;
            _Scenarios.Add(name, s);
            return s;
        }

        /// <summary>
        /// Retrieves the first company instance with the same name
        /// </summary>
        /// <param name="name"></param>
        /// <returns>Custom company instance. An exception will be thrown if not found!</returns>
        public static Company GetCompany(String name)
        {
            return _Companies[name];
        }
        /// <summary>
        /// Retrieves the first Generator instance with the same name
        /// </summary>
        /// <param name="name"></param>
        /// <returns>Custom generator instance. An exception will be thrown if not found!</returns>
        public static Generator GetGenerator(String name)
        {
            return _Generators[name];
        }

        /// <summary>
        /// Retrieves the first Reserve instance with the same name
        /// </summary>
        /// <param name="name"></param>
        /// <returns>PLEXOS reserve. An exception will be thrown if not found!</returns></returns>
        public static Reserve GetReserve(String name)
        {
            return _Reserves[name];
        }
        /// <summary>
        /// Retrieves the first Reserve instance with the same name
        /// </summary>
        /// <param name="name"></param>
        /// <returns>Custom scenario instance. </returns>
        public static Scenario GetScenario(String name)
        {
            return _Scenarios[name];
        }
        /// <summary>
        /// Function to safely check if the given element exists in this registry
        /// </summary>
        /// <param name="classEnum">PLEXOS class type</param>
        /// <param name="name">object's name</param>
        /// <returns></returns>
        public static bool Contains(EEUTILITY.Enums.ClassEnum classEnum, String name)
        {
            try
            {
                switch (classEnum)
                {
                    case EEUTILITY.Enums.ClassEnum.Generator:
                        Generator g = GetGenerator(name);
                        return true;
                    case EEUTILITY.Enums.ClassEnum.Company:
                        Company c = GetCompany(name);
                        return true;
                    case EEUTILITY.Enums.ClassEnum.Reserve:
                        Reserve r = GetReserve(name);
                        return true;
                    case EEUTILITY.Enums.ClassEnum.Scenario:
                        Scenario s = GetScenario(name);
                        return true;
                    default:
                        return false;
                }
            }
            catch (Exception e)
            {
                return false;
            }
        }

        /// <summary>
        /// Function to safely check if the given element exists in this registry
        /// </summary>
        /// <param name="classEnum"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static bool Exists(EEUTILITY.Enums.ClassEnum classEnum, String name)
        {
            switch (classEnum)
            {
                case EEUTILITY.Enums.ClassEnum.Company:
                    return _Companies.ContainsKey(name);
                case EEUTILITY.Enums.ClassEnum.Generator:
                    return _Generators.ContainsKey(name);
                case EEUTILITY.Enums.ClassEnum.Reserve:
                    return _Reserves.ContainsKey(name);
                case EEUTILITY.Enums.ClassEnum.Scenario:
                    return _Scenarios.ContainsKey(name);
                default:
                    return false;
            }
        }

        /// <summary>
        /// List of all registered Generators
        /// </summary>
        public static List<Generator> Generators
        {
            get
            {
                return new List<Generator>(_Generators.Values);
            }
        }
        /// <summary>
        /// List of all registered Reserves
        /// </summary>
        public static List<Reserve> Reserves
        {
            get
            {
                return new List<Reserve>(_Reserves.Values);
            }
        }
        /// <summary>
        /// List of all registered Companies
        /// </summary>
        public static List<Company> Companies
        {
            get
            {
                return new List<Company>(_Companies.Values);
            }
        }
        /// <summary>
        /// List of all registered Scenarios
        /// </summary>
        public static List<Scenario> Scenarios
        {
            get
            {
                return new List<Scenario>(_Scenarios.Values);
            }
        }

        private static String strSystemName;
        public static String SystemName
        {
            get => strSystemName;
            set => strSystemName = value;
        }

    }
    class Company
    {
        /// <summary>
        /// List of all generators with memberships to this company
        /// </summary>
        public List<Generator> Generators = new List<Generator>();
        private string _name;
        public String Name
        {
            get => _name;
            set => _name = value;
        }
        /// <summary>
        /// Total max response for the given company
        /// </summary>
        /// <param name="r"></param>
        /// <returns></returns>
        public double GetMaxResponse(Reserve r)
        {
            double dTotalMaxResponse = 0.0;
            foreach (Generator g in Generators)
            {
                dTotalMaxResponse += g.GetMaxResponse(r);
            }
            return dTotalMaxResponse;
        }
        /// <summary>
        /// Total max replacement for the given company
        /// </summary>
        /// <param name="r"></param>
        /// <returns></returns>
        public double GetMaxReplacement(Reserve r)
        {
            double dTotalReplacement = 0.0;
            foreach (Generator g in Generators)
            {
                dTotalReplacement += g.GetMaxResponse(r);
            }
            return dTotalReplacement;
        }
    }
    class Generator
    {
        /// <summary>
        /// List of all reserves with memberships to this generator
        /// </summary>
        public List<Reserve> Reserves = new List<Reserve>();
        private String _name;
        private Dictionary<Reserve, double> _maxresponse = new Dictionary<Reserve, double>();
        private Dictionary<Reserve, double> _maxreplacement = new Dictionary<Reserve, double>();

        public Generator(String name)
        {
            _name = name;
        }
        public String Name
        {
            get => _name;
            set => _name = value;
        }

        public double GetMaxResponse(Reserve r)
        {
            return _maxresponse[r];
        }
        public double GetMaxReplacement(Reserve r)
        {
            return _maxreplacement[r];
        }
    }

    class Reserve
    {
        /// <summary>
        /// List of all generators with memberships to this reserve
        /// </summary>
        public List<Generator> Generators = new List<Generator>();
        /// <summary>
        /// List of all scenarios that should "activate/deactivate" this reserve object
        /// </summary>
        public List<String> Scenarios = new List<String>();

        private String _name;
        private double _type = -1;
        public String Name
        {
            get => _name;
            set => _name = value;
        }
        public double Type
        {
            get => _type;
            set => _type = value;
        }
        //public String Scenario { get; set; }
    }

    class Model
    {
        public List<String> Scenarios = new List<String>();

        public String Name { get; set; }
        public String Horizon { get; set; }
        public String Report { get; set; }
        public String MT { get; set; }
        public String Transmission { get; set; }
        public String Production { get; set; }
        public String Stochastic { get; set; }
        public String Performance { get; set; }
        //public String Diagnostic { get; set; }

        public Diagnostic Diagnostic { get; set; }
    }

    class Scenario
    {
        public String Name { get; set; }

        public bool Exists { get; set; }

    }

    class Report
    {

        public List<String> Properties = new List<String>();

        public String Name { get; set; }

        public bool Summary_Year { get; set; }

    }

    class ReportProperty
    {
        public String parent_class { get; set; }
        public String child_class { get; set; }
        public String collection { get; set; }
        public String property { get; set; }

    }

    class Diagnostic
    {
        public String Name { get; set; }
        public int LPFiles { get; set; }
        public int ObjectiveFunction { get; set; }
        public int LPProgress { get; set; }
        public int MIPProgress { get; set; }
        public int Times { get; set; }
        public int StepSummary { get; set; }
        public int PerformanceSummary { get; set; }
        public int TaskSize { get; set; }
        public int DatabaseLoad { get; set; }
        public int DataFileRead { get; set; }
        public int ComputerInformation { get; set; }

    }

}
