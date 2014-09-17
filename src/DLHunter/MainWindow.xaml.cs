using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace DLHunter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private List<Outlook.AddressEntry> _seenDLs;
        private Outlook.Application _outlook;

        public MainWindow()
        {
            InitializeComponent();
            _outlook = new Outlook.Application(); // TODO: is it ok to instantiate this once?

            txtAlias.Focus();
        }

        private void btnFromAlias_Click(object sender, RoutedEventArgs e)
        {
            _seenDLs = new List<Outlook.AddressEntry>(); // TODO: this should only happpen once (perhaps at Enum with empty prefix?)
            lstMembers.Items.Clear();

            var alias = txtAlias.Text.Trim();

            // this is really crappy logic - if the user did not supply a valid email address use their's
            if (!EmailValidator.IsValid(alias))
            {
                var user = _outlook.Application.Session.CurrentUser.AddressEntry.GetExchangeUser();
                if (user != null && user.PrimarySmtpAddress != null) 
                {
                    var domain = user.PrimarySmtpAddress.Split('@')[1];
                    alias += "@" + domain;
                }
            }

            
            // Get an address entry from this alias - it looks like the easiest way to do this is to add
            // this to a mail item and then get Outlook to handle resolving it
            var item = _outlook.CreateItem(Outlook.OlItemType.olMailItem);
            
            Outlook.Recipient recipient = item.Recipients.Add(alias);
            recipient.Resolve();

            // delete the mail item
            item.Delete();

            System.Diagnostics.Debug.WriteLine(recipient.AddressEntry.Address + "   " + recipient.AddressEntry.Name);

            if (recipient != null)
            {
                var address = recipient.AddressEntry;

                if (address == null) 
                {
                    System.Diagnostics.Debug.WriteLine("not found");
                } 
                else if (address.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry)
                {
                    EnumerateDL(address);
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("Not a DL");
                }
            }

            // TODO: /facepalm :(
            txtTotalMembers.Text = "Total Members: " + lstMembers.Items.Count;
        }

        private void btnFromOutlook_Click(object sender, RoutedEventArgs e)
        {
             _seenDLs = new List<Outlook.AddressEntry>();
             lstMembers.Items.Clear();

            Outlook.SelectNamesDialog snd = _outlook.Session.GetSelectNamesDialog();
            Outlook.AddressLists addrLists = _outlook.Session.AddressLists;

            foreach (Outlook.AddressList addrList in addrLists)
            {
                if (addrList.Name == "All Groups")
                {
                    snd.InitialAddressList = addrList;
                    break;
                }
            }
            snd.NumberOfRecipientSelectors = Outlook.OlRecipientSelectors.olShowTo;
            snd.ToLabel = "D/L";
            snd.ShowOnlyInitialAddressList = true;
            snd.AllowMultipleSelection = false;
            snd.Display();

            if (snd.Recipients.Count > 0)
            {
                Outlook.AddressEntry addrEntry = snd.Recipients[1].AddressEntry;

                if (addrEntry.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry)
                {
                    EnumerateDL(addrEntry);
                }
            }

            // TODO: /facepalm :(
            txtTotalMembers.Text = "Total Members: " + lstMembers.Items.Count;
        }

        private void EnumerateDL(Outlook.AddressEntry dl, string prefix = "")
        {
            if (AlreadyEnumeratedDL(dl))
            {
                System.Diagnostics.Debug.WriteLine(prefix + dl.Name + ": <skipping nested dl>");
            }
            else
            {
                _seenDLs.Add(dl);
                
                Outlook.ExchangeDistributionList exchDL = dl.GetExchangeDistributionList();
                Outlook.AddressEntries addrEntries = exchDL.GetExchangeDistributionListMembers();

                if (addrEntries != null)
                {
                    foreach (Outlook.AddressEntry exchDLMember in addrEntries)
                    {
                        if (exchDLMember.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry)
                        {
                            EnumerateDL(exchDLMember, prefix + dl.Name + ": ");
                        }
                        else
                        {
                            var entry = prefix + dl.Name + ": " + exchDLMember.Name;
                            System.Diagnostics.Debug.WriteLine(entry);
                            lstMembers.Items.Add(entry);
                        }
                    }
                }
            }
        }

        private bool AlreadyEnumeratedDL(Outlook.AddressEntry target)
        {
            foreach (var dl in _seenDLs)
            {
                if (dl.Address == target.Address)
                    return true;
            }

            return false;
        }

        private void txtAlias_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                btnFromAlias_Click(null, null);
        }

    }
}
