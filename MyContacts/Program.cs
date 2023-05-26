using CsvHelper;
using CsvHelper.Configuration.Attributes;
using HtmlAgilityPack;
using HtmlAgilityPack.CssSelectors.NetCore;
using PhoneNumbers;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace MyContacts
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var contactsGroup = ContactsHelper.Split(ContactsHelper.LoadGhanaYello(1, totalPages: 753, "Accra"), chunkSize: 1000);
            contactsGroup.ForEach(contacts =>
            {
                ContactsHelper.Save(contacts, $"./GhanaYello Accra Contacts {contactsGroup.IndexOf(contacts) + 1}.csv");
            });
        }
    }

    public class ContactsHelper
    {
        public class Contact
        {
            [Name("Name")]
            public string Name { get; set; }

            [Name("Given Name")]
            public string GivenName { get; set; }

            [Name("Additional Name")]
            public string AdditionalName { get; set; }

            [Name("Family Name")]
            public string FamilyName { get; set; }

            [Name("E-mail 1 - Type")]
            public string EmailType1 { get; set; } = "* Office";

            [Name("E-mail 1 - Value")]
            public string EmailValue1 { get; set; }

            [Name("Phone 1 - Type")]
            public string PhoneType1 { get; set; } = "Work";

            [Name("Phone 1 - Value")]
            public string PhoneValue1 { get; set; }

            [Name("Organization 1 - Name")]
            public string OrganizationName1 { get; set; }

            [Name("Organization 1 - Title")]
            public string OrganizationTitle1 { get; set; }

            [Name("Organization 1 - Department")]
            public string OrganizationDepartment1 { get; set; }

            [Name("Notes")]
            public string Notes { get; set; }

            [Name("Website 1 - Value")]
            public string WebsiteValue1 { get; set; }

            [Name("Birthday")]
            public string Birthday { get; set; }

            [Name("Gender")]
            public string Gender { get; set; }

            [Name("Location")]
            public string Location { get; set; }

            [Name("Address 1 - Formatted")]
            public string AddressFormatted1 { get; set; }

            [Name("Address 1 - Street")]
            public string AddressStreet1 { get; set; }

            [Name("Address 1 - City")]
            public string AddressCity1 { get; set; }

            [Name("Address 1 - PO Box")]
            public string AddressPOBox1 { get; set; }

            [Name("Address 1 - Region")]
            public string AddressRegion1 { get; set; }

            [Name("Address 1 - Postal Code")]
            public string AddressPostalCode1 { get; set; }

            [Name("Address 1 - Country")]
            public string AddressCountry1 { get; set; }

            [Name("Address 1 - Extended Address")]
            public string AddressExtended1 { get; set; }

            [Name("Organization 1 - Type")]
            public string OrganizationType1 { get; set; }

            [Name("Billing Information")]
            public string BillingInformation { get; set; }

            [Name("Directory Server")]
            public string DirectoryServer { get; set; }

            [Name("Mileage")]
            public string Mileage { get; set; }

            [Name("Occupation")]
            public string Occupation { get; set; }

            [Name("Hobby")]
            public string Hobby { get; set; }

            [Name("Initials")]
            public string Initials { get; set; }

            [Name("Nickname")]
            public string Nickname { get; set; }

            [Name("Short Name")]
            public string ShortName { get; set; }

            [Name("Maiden Name")]
            public string MaidenName { get; set; }

            [Name("Sensitivity")]
            public string Sensitivity { get; set; }

            [Name("Name Suffix")]
            public string NameSuffix { get; set; }

            [Name("Name Prefix")]
            public string NamePrefix { get; set; }

            [Name("Priority")]
            public string Priority { get; set; }

            [Name("Subject")]
            public string Subject { get; set; }

            [Name("Language")]
            public string Language { get; set; }

            [Name("Photo")]
            public string Photo { get; set; }

            [Name("Group Membership")]
            public string GroupMembership { get; set; }

            [Name("Address 1 - Type")]
            public string AddressType1 { get; set; }

            [Name("Organization 1 - Symbol")]
            public string OrganizationSymbol1 { get; set; }

            [Name("Organization 1 - Location")]
            public string OrganizationLocation1 { get; set; }

            [Name("Organization 1 - Job Description")]
            public string OrganizationJobDescription1 { get; set; }

            [Name("Website 1 - Type")]
            public string WebsiteType1 { get; set; }
        }

        public static Contact[] LoadGhanaYello(int initialPage, int totalPages, string location)
        {
            var companies = new List<Contact>();

            var baseUrl = "https://www.ghanayello.com";

            for (int currentPage = (initialPage - 1); currentPage < totalPages; currentPage++)
            {
                var listPageUrl = $"{baseUrl}/location/{location}/1";
                var listPageNodes = HtmlHelper.LoadDocument(listPageUrl).DocumentNode.QuerySelectorAll("#listings > .company:not(.company_ad)");

                foreach (var listPageNode in listPageNodes)
                {
                    var pageUrl = $"{baseUrl}/{listPageNode.QuerySelector("h4 > a").GetAttributeValue("href", string.Empty).TrimStart('/')}";
                    var pageNode = HtmlHelper.LoadDocument(pageUrl).QuerySelector("#company_item");

                    var phoneNumbers = HtmlHelper.ValidPhones(HtmlHelper.StripHtml(pageNode.QuerySelector(".info > .label ~ .phone").InnerHtml));
                    var companyLocation = HtmlHelper.StripHtml(string.Join("", pageNode.QuerySelector(".info > .label ~ .location").ChildNodes.Take(3).Select(_ => _.InnerText)));
                    var companyName = HtmlHelper.StripHtml(pageNode.QuerySelector(".info > .label ~ #company_name").InnerText);

                    foreach (var companyPhoneNumber in phoneNumbers)
                    {
                        companies.Add(new Contact { Name = companyName, PhoneValue1 = companyPhoneNumber, WebsiteValue1 = pageUrl, Location = companyLocation });

                        Console.WriteLine($"{companies.Count}. Phone number: {companyPhoneNumber}, Name: {companyName}, Location: {companyLocation}");
                    }
                }
            }

            return companies.ToArray();
        }

        public static void Save(Contact[] array, string path)
        {
            using var writer = new StreamWriter(new FileStream(path, FileMode.Create));
            using var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);
            csv.WriteRecords(array);
        }

        public static List<Contact[]> Split(Contact[] array, int chunkSize)
        {
            List<Contact[]> chunks = new List<Contact[]>();

            for (int i = 0; i < array.Length; i += chunkSize)
            {
                int size = Math.Min(chunkSize, array.Length - i);
                Contact[] chunk = new Contact[size];
                Array.Copy(array, i, chunk, 0, size);
                chunks.Add(chunk);
            }

            return chunks;
        }

        public static class HtmlHelper
        {
            public static HtmlDocument LoadDocument(string url)
            {
                try
                {
                    return new HtmlWeb().Load(url);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                    Console.WriteLine("Retrying in 5 seconds...");
                    Thread.Sleep(5000);

                    return LoadDocument(url);
                }
            }
            public static string StripHtml(string html)
            {
                var doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(html);

                if (doc.DocumentNode == null || doc.DocumentNode.ChildNodes == null)
                {
                    return WebUtility.HtmlDecode(html);
                }

                var sb = new StringBuilder();
                var i = 0;

                foreach (var node in doc.DocumentNode.ChildNodes)
                {
                    var text = node.InnerText?.Trim();

                    if (!string.IsNullOrEmpty(text))
                    {
                        sb.Append(text);

                        if (i < doc.DocumentNode.ChildNodes.Count - 1)
                        {
                            sb.Append(Environment.NewLine);
                        }
                    }

                    i++;
                }

                return WebUtility.HtmlDecode(sb.ToString());
            }

            public static string WrapHtml(string text)
            {
                text = WebUtility.HtmlEncode(text);
                text = text.Replace("\r\n", "\r");
                text = text.Replace("\n", "\r");
                text = text.Replace("\r", "<br>\r\n");
                text = text.Replace("  ", " &nbsp;");
                return $"<p>{text.Trim()}</p>";
            }

            public static string[] ValidPhones(string number)
            {
                return InternalValidPhones(number).Distinct().ToArray();
            }

            public static IEnumerable<string> InternalValidPhones(string number)
            {
                var matches = Regex.Matches(number.Replace(Environment.NewLine, " and "), "(\\+*)(\\-*\\(*\\s*\\d\\s*\\)*\\-*){9,}");

                foreach (var match in matches)
                {
                    var phoneNumber = match.ToString()!.Replace("(", "")!.Replace(")", "").Replace("+", "").Replace("-", "").Replace(" ", "").TrimStart("00");

                    var phoneUtil = PhoneNumberUtil.GetInstance();
                    var phoneNumberObj = phoneUtil.Parse(phoneNumber, "GH");

                    if (phoneUtil.IsValidNumber(phoneNumberObj))
                    {
                        phoneNumber = phoneUtil.Format(phoneNumberObj, PhoneNumberFormat.NATIONAL);
                        yield return phoneNumber;
                    }
                }
            }
        }
    }


    public static class Extensions
    {
        public static string TrimStart(this string source, string value, StringComparison comparisonType = StringComparison.Ordinal)
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            int valueLength = value.Length;
            int startIndex = 0;
            while (source.IndexOf(value, startIndex, comparisonType) == startIndex)
            {
                startIndex += valueLength;
            }

            return source.Substring(startIndex);
        }

        public static string TrimEnd(this string source, string value, StringComparison comparisonType = StringComparison.Ordinal)
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            int sourceLength = source.Length;
            int valueLength = value.Length;
            int count = sourceLength;
            while (source.LastIndexOf(value, count, comparisonType) == count - valueLength)
            {
                count -= valueLength;
            }

            return source.Substring(0, count);
        }

    }
}