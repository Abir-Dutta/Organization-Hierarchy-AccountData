using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.IO;
using System.Linq;

namespace FetchData
{
    // Program to get Organization Hierarchy from LADP server
    public static class Program
    {
        static string path = @"c:\temp\MyTest.txt";
        const int AllLevel = -9;
        static List<string> strList = new List<string>();
        static int spacerCount = 0;
        static List<Domain> domainList = Forest.GetCurrentForest().Domains.Cast<Domain>().ToList();
        // EMP = Employee | CWR = Contractor
        static List<string> employeeTypeList = new List<string>() { "EMP", "CWR" };

        static void CreateFile() {
            Directory.CreateDirectory(string.Join("\\", path.Split('\\').TakeWhile(p => !p.Contains(".txt"))));
            File.Create(path).Close();
        }

        static void WritetoFile(string writeTheLine)
        {
            File.AppendAllText(path, writeTheLine);
        }
        static string CutName(string Lpath)
        {
            return Lpath?.Replace("LDAP://", "").Replace("\\,", "*").Replace("/", ",").Split(',')
                    .FirstOrDefault(x => x.StartsWith("CN="))?.Replace("CN=", "").Replace("*", ",");
        }
        static string CutLdapPath(string Lpath)
        {
            return string.Join(",", Lpath?.Replace("LDAP://", "").Split(',')
                    .Where(x => x.StartsWith("DC="))?.ToList());
        }
        static string Spacer(int count)
        {
            string spacer = "";
            for (int i = 0; i < count; i++)
            {
                spacer += "      ";
            }
            return spacer;
        }

        static string GetMemberOf(ResultPropertyValueCollection memberOfCollection)
        {
            List<string> memberOfList = new List<string>();
            foreach ( var item in memberOfCollection)
            {
                memberOfList.Add(CutName(Convert.ToString(item)));
            }
            memberOfList.Sort();

            return string.Join(", ",memberOfList);
        }
        static void FindUser(string userName, string managerName, int level = AllLevel)
        {
            foreach (Domain domain in domainList)
            {
                try
                {
                    DirectorySearcher searcher = new DirectorySearcher();
                    SearchResultCollection results = null;

                    //Other filters can be applied cn,displayname, etc.
                    //Based on the LDAP properties need to change the below Properties to Load. 
                    searcher.Filter = "(&(objectClass=user)(objectCategory=person)(cn=" + userName + "))";
                    //Common Name
                    searcher.PropertiesToLoad.Add("cn");  
                    searcher.PropertiesToLoad.Add("displayname");
                    searcher.PropertiesToLoad.Add("directreports");
                    searcher.PropertiesToLoad.Add("userprincipalname");
                    //searcher.PropertiesToLoad.Add("samaccountname");
                    searcher.PropertiesToLoad.Add("department");
                    searcher.PropertiesToLoad.Add("title");
                    //Location
                    searcher.PropertiesToLoad.Add("l");
                    searcher.PropertiesToLoad.Add("co");
                    searcher.PropertiesToLoad.Add("manager");
                    searcher.PropertiesToLoad.Add("employeetype");
                    searcher.PropertiesToLoad.Add("msexchhidefromaddresslists");
                    searcher.PropertiesToLoad.Add("memberof");
                    DirectoryEntry entry = new DirectoryEntry("LDAP://" + domain.Name);
                    searcher.SearchRoot = entry;
                    searcher.SearchScope = SearchScope.Subtree;

                    results = searcher.FindAll();


                    if (results.Count == 0) continue;

                    foreach (SearchResult result in results)
                    {
                        string manager = CutName(Convert.ToString(result.Properties["manager"][0]));


                        if (manager.Equals(managerName)
                            && result.Properties["userprincipalname"].Count > 0
                            && (result.Properties["msexchhidefromaddresslists"].Count > 0 ? !Convert.ToBoolean(result.Properties["msexchhidefromaddresslists"][0]) : true)
                            && (result.Properties["employeetype"].Count > 0 && employeeTypeList.Contains(Convert.ToString(result.Properties["employeetype"][0]))))
                        {
                            string name = CutName(Convert.ToString(result.Path));

                            string finalTitle = result.Properties["title"].Count > 0 ? Convert.ToString(result.Properties["title"][0]) : (result.Properties["employeetype"].Count > 0 && Convert.ToString(result.Properties["employeetype"][0]).Equals("CWR") ? "CONTRACTOR" : "");
                            if (string.IsNullOrEmpty(finalTitle)) continue;
                            strList.Add(spacerCount.ToString()); // Level
                            strList.Add(Spacer(spacerCount) + (result.Properties["displayname"].Count > 0 ? Convert.ToString(result.Properties["displayname"][0]).Split('@').First() : ""));
                            strList.Add(finalTitle);
                            strList.Add(result.Properties["directreports"].Count > 0 ? "" : "0");
                            strList.Add(result.Properties["l"].Count > 0 ? Convert.ToString(result.Properties["l"][0]) : "");
                            strList.Add(result.Properties["co"].Count > 0 ? Convert.ToString(result.Properties["co"][0]) : "");
                            strList.Add(result.Properties["department"].Count > 0 ? Convert.ToString(result.Properties["department"][0]) : "");
                            strList.Add(managerName.Split('@').First());
                            strList.Add(result.Properties["employeetype"].Count > 0 && Convert.ToString(result.Properties["employeetype"][0]).Equals("CWR") ? "Contractor" : "FTE");
                            strList.Add(result.Properties["userprincipalname"].Count > 0 ? Convert.ToString(result.Properties["userprincipalname"][0]) : "");
                            strList.Add(result.Properties["memberof"].Count > 0 ? GetMemberOf(result.Properties["memberof"]) : "");

                            strList.Add(Environment.NewLine);

                            WritetoFile(string.Join("|", strList));
                            strList.Clear();
                            spacerCount++;

                            if(level != AllLevel) level--;

                            if (level >= 0 || level == AllLevel)
                            {
                                foreach (var reportee in result.Properties["directreports"])
                                {
                                    Console.WriteLine(reportee);
                                    // ignore Service/Robot accounts
                                    if (CutLdapPath(Convert.ToString(reportee)).ToLower().Contains("dc=tnd,")) continue;
                                    FindUser(CutName(Convert.ToString(reportee)), name, level);
                                }
                            }
                            spacerCount--;

                            break;
                        }
                    }
                    break;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

            }

        }
        static void Main(string[] args)
        {
            CreateFile();

            //Pass filter values based on the filter type. cn,displayname, etc.
            //FindUser("Dutta, Abir @ Contractor", "managerlastname, managerfirstname @ India");
            FindUser("lastname, firstname @ location", "managerlastname, managerfirstname @ location");
        }

    }

}
