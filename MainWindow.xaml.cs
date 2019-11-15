using System;
using System.Windows;
using Microsoft.Office.Interop.Outlook;

namespace LunchBot
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {

        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            // google form for Curry


            // outlook

            Microsoft.Office.Interop.Outlook.Application otApp = new Outlook.Application();// create outlook object
            Outlook.MailItem otMsg = (Outlook.MailItem)otApp.CreateItem(Outlook.OlItemType.olMailItem); // Create mail object
            Outlook.Recipient otRecip = (Outlook.Recipient)otMsg.Recipients.Add(EmailUrl);
            otRecip.Resolve();// validate recipient address
            otMsg.Subject = "Test Subject";
            otMsg.Body = "Text Message";
            String sSource = AppDomain.CurrentDomain.BaseDirectory + "Test.txt";

        }
    }
}
