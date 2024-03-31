using System;
using System.IO;
using System.Net;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Net.Sockets;
using Microsoft.Office.Interop.Outlook;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace OutlookAddIn2
{

    public partial class ThisAddIn
    {

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.ItemSend += Application_ItemSend;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //注: Outlook はこのイベントを発行しなくなりました。Outlook が
            //    を Outlook のシャットダウン時に実行する必要があります。https://go.microsoft.com/fwlink/?LinkId=506785 をご覧ください
        }


        private void Application_ItemSend(object Item, ref bool Cancel)
        {

            // 送信ボタンが押されたときの処理を記述します。
            Outlook.MailItem mailItem = Item as Outlook.MailItem;
            if (mailItem != null)
            {

                // 送信されるメールの件名を取得します。
                string subject = mailItem.Subject;

                // 送信されるメールの本文を取得します。
                string body = mailItem.Body;

                // 送信されるメールの差出人を取得します。
                string senderEmailAddress = mailItem.SenderEmailAddress;

                // 送信されるメールの宛先を取得します。
                StringBuilder recipients = new StringBuilder();
                foreach (Outlook.Recipient recipient in mailItem.Recipients)
                {
                    recipients.AppendLine("Recipient Name: " + recipient.Name);
                    recipients.AppendLine("Recipient Address: " + recipient.Address);
                    recipients.AppendLine("Recipient Type: " + recipient.Type);
                }

                // 送信されるメールの添付ファイルを取得します。
                StringBuilder attachmentsInfo = new StringBuilder();
                foreach (Outlook.Attachment attachment in mailItem.Attachments)
                {
                    attachmentsInfo.AppendLine("Attachment Name: " + attachment.DisplayName);
                    attachmentsInfo.AppendLine("Attachment Type: " + attachment.Type);
                    attachmentsInfo.AppendLine("Attachment Size: " + attachment.Size + " bytes");
                }

                // 件名と本文をコンソールに表示します。必要に応じて他の処理を行います。
                //MessageBox.Show("senderEmailAddress: " + senderEmailAddress);
                //MessageBox.Show("recipients: " + recipients.ToString());
                //MessageBox.Show("Subject: " + subject);
                //MessageBox.Show("Body: " + body);
                //MessageBox.Show("recipients: " + attachmentsInfo.ToString());

                string url = "http://localhost/test/post/post.php";

                System.Net.WebClient wc = new System.Net.WebClient();
                //NameValueCollectionの作成
                System.Collections.Specialized.NameValueCollection ps =
                    new System.Collections.Specialized.NameValueCollection();
                //送信するデータ（フィールド名と値の組み合わせ）を追加
                ps.Add("word", "インターネット");
                ps.Add("id", "1");
                //データを送信し、また受信する
                byte[] resData = wc.UploadValues(url, ps);
                wc.Dispose();

                //受信したデータを表示する
                string resText = System.Text.Encoding.UTF8.GetString(resData);
                MessageBox.Show(resText);

                // ここにブラウザを開く処理を記述します。
                //string url = "https://www.example.com";
                //System.Diagnostics.Process.Start(url);

                Cancel = true;

            }

        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }

    
}
