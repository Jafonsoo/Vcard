using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Vcard.Models;

namespace Vcard.Helpers
{
    public class FileHelper
    {
        const string NewLine = "\r\n";
        const string Separator = ";";
        const string Header = "BEGIN:VCARD\r\nVERSION:2.1";
        const string Name = "N:";
        const string FormattedName = "FN:";
        const string OrganizationName = "ORG:";
        const string TitlePrefix = "TITLE:";
        const string PhotoPrefix = "PHOTO;ENCODING=BASE64;JPEG:";
        const string PhonePrefix = "TEL;WORK;VOICE:";
        const string PhoneSubPrefix = ",VOICE:";
        const string AddressPrefix = "ADR;TYPE=";
        const string AddressSubPrefix = ":;;";
        const string EmailPrefix = "EMAIL:";
        const string Footer = "END:VCARD";
        const string Title = "TITLE:";

        public static string CreateVCard(User contact)
        {
            StringBuilder fw = new StringBuilder();
            fw.Append(Header);
            fw.Append(NewLine);

            //Full Name
            if (!string.IsNullOrEmpty(contact.FirstName) || !string.IsNullOrEmpty(contact.LastName))
            {
                fw.Append(Name);
                fw.Append(contact.LastName);
                fw.Append(Separator);
                fw.Append(contact.FirstName);
                fw.Append(Separator);
                fw.Append(NewLine);
            }

            //Formatted Name
            if (!string.IsNullOrEmpty(contact.OtherName))
            {
                fw.Append(FormattedName);
                fw.Append(contact.OtherName);
                fw.Append(NewLine);
            }

            //Organization name
            if (!string.IsNullOrEmpty(contact.Company))
            {
                fw.Append(OrganizationName);
                fw.Append(contact.Company);
                fw.Append(";");
                fw.Append(contact.Departamento);
                fw.Append(NewLine);
            }

            //Phones
            if (!string.IsNullOrEmpty(contact.Phone))
            {
                fw.Append(PhonePrefix);
                fw.Append(contact.Phone);
                fw.Append(NewLine);
            }

            //Email
            if (!string.IsNullOrEmpty(contact.EmailAddress))
            {
                fw.Append(EmailPrefix);
                fw.Append(contact.EmailAddress);
                fw.Append(NewLine);
            }

            if (!string.IsNullOrEmpty(contact.Title))
            {
                fw.Append(Title);
                fw.Append(contact.Title);
                fw.Append(NewLine);
            }

            fw.Append(Footer);

            return fw.ToString();
        }
    }
}
