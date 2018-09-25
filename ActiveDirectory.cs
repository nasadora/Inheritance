using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices.ActiveDirectory;
using System.DirectoryServices;
using System.Text.RegularExpressions;
using System.Data;
using System.IO;

namespace ActiveDirectory
{
    public class Active_Directory
    {
        public string adAttribute;
        public string getAdInfo(string adMID, string userDomain, string adACT)
        {           
           
            string DomainPath = "LDAP://" + userDomain + "/DC=" + userDomain + ",DC=*****,DC=****";
            using (DirectoryEntry stringDomain = new DirectoryEntry(DomainPath)) // create the string for the domain
            {
                using (DirectorySearcher stringSearch = new DirectorySearcher(stringDomain)) //creating the search on the string
                {
                    stringSearch.Filter = "(samaccountname=" + adMID + ")"; // creates the filter to be cn= then the mud id
                    SearchResult result = stringSearch.FindOne(); // looks in the adStringDomain for the user and finds the one that will match

                    DirectoryEntry adUserAttribute = result.GetDirectoryEntry(); // now use the search to look for attributes
                    int flags = (int)adUserAttribute.Properties["userAccountControl"].Value; // convert user account control from string to Int
                    if (!Convert.ToBoolean(flags & 0x0002) == false) // if does not divide its disabled  
                    {
                        adAttribute = "N";
                    }
                    else
                    {
                        adAttribute = "Y";
                    }
                }
            }

            return adAttribute;
        }

    }
}

