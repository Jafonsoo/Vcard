// See https://aka.ms/new-console-template for more information
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Security.Claims;
using System.Security.Principal;
using Vcard.Helpers;
using Vcard.Models;

string root = "aquinos.org";
List<User> adAllUsers_ = new List<User>();

using (PrincipalContext pc = new PrincipalContext(ContextType.Domain, root))
{

    using (var searcher = new PrincipalSearcher(new UserPrincipal(pc)))
    {
        foreach (var result in searcher.FindAll())
        {
            DirectoryEntry de = result.GetUnderlyingObject() as DirectoryEntry;

            if (de.Properties["givenName"].Value != null &&
                de.Properties["sn"].Value != null &&
                de.Properties["mail"].Value != null &&
                de.Properties["cn"].Value != null &&
                de.Properties["telephoneNumber"].Value != null &&
                de.Properties["company"].Value != null &&
                de.Properties["department"].Value != null &&
                de.Properties["title"].Value != null)
            {
                adAllUsers_.Add(new User
                {
                    FirstName = de.Properties["givenName"].Value.ToString(),
                    LastName = de.Properties["sn"].Value.ToString(),
                    OtherName = de.Properties["cn"].Value.ToString(),
                    EmailAddress = de.Properties["mail"].Value.ToString(),
                    Phone = de.Properties["telephoneNumber"].Value.ToString(),
                    Departamento = de.Properties["department"].Value.ToString(),
                    Company = de.Properties["company"].Value.ToString(),
                    Title = de.Properties["title"].Value.ToString()
                });
            }
            else if (de.Properties["givenName"].Value != null &&
                    de.Properties["sn"].Value != null &&
                    de.Properties["mail"].Value != null &&
                    de.Properties["cn"].Value != null &&
                    de.Properties["telephoneNumber"].Value != null &&
                    de.Properties["company"].Value != null &&
                    de.Properties["title"].Value != null)
            {
                adAllUsers_.Add(new User
                {
                    FirstName = de.Properties["givenName"].Value.ToString(),
                    LastName = de.Properties["sn"].Value.ToString(),
                    OtherName = de.Properties["cn"].Value.ToString(),
                    EmailAddress = de.Properties["mail"].Value.ToString(),
                    Phone = de.Properties["telephoneNumber"].Value.ToString(),
                    Company = de.Properties["company"].Value.ToString(),
                    Title = de.Properties["title"].Value.ToString()
                });
            }
            else if (de.Properties["givenName"].Value != null &&
                    de.Properties["sn"].Value != null &&
                    de.Properties["mail"].Value != null &&
                    de.Properties["cn"].Value != null &&
                    de.Properties["telephoneNumber"].Value != null &&
                    de.Properties["department"].Value != null &&
                    de.Properties["company"].Value != null)
            {
                adAllUsers_.Add(new User
                {
                    FirstName = de.Properties["givenName"].Value.ToString(),
                    LastName = de.Properties["sn"].Value.ToString(),
                    OtherName = de.Properties["cn"].Value.ToString(),
                    EmailAddress = de.Properties["mail"].Value.ToString(),
                    Phone = de.Properties["telephoneNumber"].Value.ToString(),
                    Departamento = de.Properties["department"].Value.ToString(),
                    Company = de.Properties["company"].Value.ToString()
                });
            }
            else if (de.Properties["givenName"].Value != null &&
                    de.Properties["sn"].Value != null &&
                    de.Properties["mail"].Value != null &&
                    de.Properties["cn"].Value != null &&
                    de.Properties["telephoneNumber"].Value != null &&
                    de.Properties["company"].Value != null)
            {
                adAllUsers_.Add(new User
                {
                    FirstName = de.Properties["givenName"].Value.ToString(),
                    LastName = de.Properties["sn"].Value.ToString(),
                    OtherName = de.Properties["cn"].Value.ToString(),
                    EmailAddress = de.Properties["mail"].Value.ToString(),
                    Phone = de.Properties["telephoneNumber"].Value.ToString(),
                    Company = de.Properties["company"].Value.ToString()
                });
            }
            else if (de.Properties["givenName"].Value != null &&
                     de.Properties["sn"].Value != null &&
                     de.Properties["mail"].Value != null &&
                     de.Properties["title"].Value != null &&
                     de.Properties["department"].Value != null &&
                     de.Properties["telephoneNumber"].Value != null)
            {
                adAllUsers_.Add(new User
                {
                    FirstName = de.Properties["givenName"].Value.ToString(),
                    LastName = de.Properties["sn"].Value.ToString(),
                    OtherName = de.Properties["cn"].Value.ToString(),
                    EmailAddress = de.Properties["mail"].Value.ToString(),
                    Title = de.Properties["title"].Value.ToString(),
                    Departamento = de.Properties["department"].Value.ToString(),
                    Phone = de.Properties["telephoneNumber"].Value.ToString()
                });
            }
            else if (de.Properties["givenName"].Value != null &&
                     de.Properties["sn"].Value != null &&
                     de.Properties["mail"].Value != null &&
                     de.Properties["title"].Value != null &&
                     de.Properties["telephoneNumber"].Value != null)
            {
                adAllUsers_.Add(new User
                {
                    FirstName = de.Properties["givenName"].Value.ToString(),
                    LastName = de.Properties["sn"].Value.ToString(),
                    OtherName = de.Properties["cn"].Value.ToString(),
                    EmailAddress = de.Properties["mail"].Value.ToString(),
                    Title = de.Properties["title"].Value.ToString(),
                    Phone = de.Properties["telephoneNumber"].Value.ToString()
                });
            }
            else if (de.Properties["givenName"].Value != null &&
                    de.Properties["sn"].Value != null &&
                    de.Properties["mail"].Value != null &&
                    de.Properties["department"].Value != null &&
                    de.Properties["telephoneNumber"].Value != null)
            {
                adAllUsers_.Add(new User
                {
                    FirstName = de.Properties["givenName"].Value.ToString(),
                    LastName = de.Properties["sn"].Value.ToString(),
                    OtherName = de.Properties["cn"].Value.ToString(),
                    EmailAddress = de.Properties["mail"].Value.ToString(),
                    Departamento = de.Properties["department"].Value.ToString(),
                    Phone = de.Properties["telephoneNumber"].Value.ToString()
                });
            }
            else if (de.Properties["givenName"].Value != null &&
                de.Properties["sn"].Value != null &&
                de.Properties["mail"].Value != null &&
                de.Properties["telephoneNumber"].Value != null)
            {
                adAllUsers_.Add(new User
                {
                    FirstName = de.Properties["givenName"].Value.ToString(),
                    LastName = de.Properties["sn"].Value.ToString(),
                    OtherName = de.Properties["cn"].Value.ToString(),
                    EmailAddress = de.Properties["mail"].Value.ToString(),
                    Phone = de.Properties["telephoneNumber"].Value.ToString()
                });
            }
            else if (de.Properties["givenName"].Value != null &&
                    de.Properties["sn"].Value != null &&
                    de.Properties["telephoneNumber"].Value != null &&
                    de.Properties["title"].Value != null &&
                    de.Properties["department"].Value != null &&
                    de.Properties["cn"].Value != null)
            {
                adAllUsers_.Add(new User
                {
                    FirstName = de.Properties["givenName"].Value.ToString(),
                    LastName = de.Properties["sn"].Value.ToString(),
                    OtherName = de.Properties["cn"].Value.ToString(),
                    Departamento = de.Properties["department"].Value.ToString(),
                    Title = de.Properties["title"].Value.ToString(),
                    Phone = de.Properties["telephoneNumber"].Value.ToString()
                });
            }
            else if (de.Properties["givenName"].Value != null &&
                    de.Properties["sn"].Value != null &&
                    de.Properties["telephoneNumber"].Value != null &&
                    de.Properties["title"].Value != null &&
                    de.Properties["cn"].Value != null)
            {
                adAllUsers_.Add(new User
                {
                    FirstName = de.Properties["givenName"].Value.ToString(),
                    LastName = de.Properties["sn"].Value.ToString(),
                    OtherName = de.Properties["cn"].Value.ToString(),
                    Title = de.Properties["title"].Value.ToString(),
                    Phone = de.Properties["telephoneNumber"].Value.ToString()
                });
            }
            else if (de.Properties["givenName"].Value != null &&
                    de.Properties["sn"].Value != null &&
                    de.Properties["telephoneNumber"].Value != null &&
                    de.Properties["department"].Value != null &&
                    de.Properties["cn"].Value != null)
            {
                adAllUsers_.Add(new User
                {
                    FirstName = de.Properties["givenName"].Value.ToString(),
                    LastName = de.Properties["sn"].Value.ToString(),
                    OtherName = de.Properties["cn"].Value.ToString(),
                    Departamento = de.Properties["department"].Value.ToString(),
                    Phone = de.Properties["telephoneNumber"].Value.ToString()
                });
            }
            else if (de.Properties["givenName"].Value != null &&
                de.Properties["sn"].Value != null &&
                de.Properties["telephoneNumber"].Value != null &&
                de.Properties["cn"].Value != null)
            {
                adAllUsers_.Add(new User
                {
                    FirstName = de.Properties["givenName"].Value.ToString(),
                    LastName = de.Properties["sn"].Value.ToString(),
                    OtherName = de.Properties["cn"].Value.ToString(),
                    Phone = de.Properties["telephoneNumber"].Value.ToString()
                });
            }
            else if (de.Properties["telephoneNumber"].Value != null &&
                    de.Properties["cn"].Value != null)
            {
                adAllUsers_.Add(new User
                {
                    OtherName = de.Properties["cn"].Value.ToString(),
                    Phone = de.Properties["telephoneNumber"].Value.ToString()
                });
            }
        }
    }
}

List<string> adlist = new List<string>();

foreach (var ad in adAllUsers_)
{
    string vcardcontents = FileHelper.CreateVCard(ad);

    adlist.Add(vcardcontents);
}

string SavePath = System.IO.Path.Combine(AppContext.BaseDirectory, "output.vcf");
System.IO.File.WriteAllText(SavePath, String.Join("\r\n", adlist.ToArray()));

Console.WriteLine("File saved at " + SavePath.Trim());