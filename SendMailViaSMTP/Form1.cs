using System;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Net;
using System.Net.Mail;

namespace SendMailViaSMTP
{
    public partial class SendMailForm : Form
    {
        public SendMailForm()
        {
            InitializeComponent();
        }
        
        private void SendMailForm_Load(object sender, EventArgs e)
        {
            SqlConnection sqlconnection = new SqlConnection();
            sqlconnection.ConnectionString = "server =IMRAN-PC; database =EmployeeDB; User ID =sa; Password =$Nextservices";
            SqlCommand sqlCommand = new SqlCommand("select FirstName,LastName,Gender,Salary from Employees", sqlconnection); //select query command  
            SqlDataAdapter sqlDataAdapter = new System.Data.SqlClient.SqlDataAdapter();
            sqlDataAdapter.SelectCommand = sqlCommand; //add selected rows to sql data adapter  
            DataSet dataSetEmployee = new DataSet(); //create new data set  

            try
            {

                sqlDataAdapter.Fill(dataSetEmployee, "employee"); //fill sql data adapter rows to data set  
                dgvEmployee.ColumnCount = 4;
                dgvEmployee.Columns[0].HeaderText = "First Name";
                dgvEmployee.Columns[0].DataPropertyName = "FirstName";
                dgvEmployee.Columns[1].HeaderText = "Last Name";
                dgvEmployee.Columns[1].DataPropertyName = "LastName";
                dgvEmployee.Columns[2].HeaderText = "Gender";
                dgvEmployee.Columns[2].DataPropertyName = "Gender";
                dgvEmployee.Columns[3].HeaderText = "Salary";
                dgvEmployee.Columns[3].DataPropertyName = "Salary";
                dgvEmployee.DataSource = dataSetEmployee;
                dgvEmployee.DataMember = "employee";
            }
            catch (Exception Ex)
            {
                System.Windows.Forms.MessageBox.Show(Ex.Message);
                sqlconnection.Close();
            }
        }


        public static string getHtml(DataGridView grid)
        {
            try
            {
                StringBuilder sbmessageBody = new StringBuilder();
                sbmessageBody.Append("<font>The following are the records: </font><br><br>");
                if (grid.RowCount == 0) return Convert.ToString(sbmessageBody);
                string htmlTableStart = "<table style=\"border-collapse:collapse; text-align:center;\" >";
                string htmlTableEnd = "</table>";
                string htmlHeaderRowStart = "<tr style=\"background-color:#6FA1D2; color:#ffffff;\">";
                string htmlHeaderRowEnd = "</tr>";
                string htmlTrStart = "<tr style=\"color:#555555;\">";
                string htmlTrEnd = "</tr>";
                string htmlTdStart = "<td style=\" border-color:#5c87b2; border-style:solid; border-width:thin; padding: 5px;\">";
                string htmlTdEnd = "</td>";
                sbmessageBody.Append(htmlTableStart);
                sbmessageBody.Append(htmlHeaderRowStart);
                sbmessageBody.Append(htmlTdStart + "First Name" + htmlTdEnd);
                sbmessageBody.Append(htmlTdStart + "Last Name" + htmlTdEnd);
                sbmessageBody.Append(htmlTdStart + "Gender" + htmlTdEnd);
                sbmessageBody.Append(htmlTdStart + "Salary" + htmlTdEnd);
                sbmessageBody.Append(htmlHeaderRowEnd);
                //Loop all the rows from grid vew and added to html td  
                for (int i = 0; i <= grid.RowCount - 1; i++)
                {
                    sbmessageBody.Append(htmlTrStart);
                    sbmessageBody.Append(htmlTdStart + grid.Rows[i].Cells[0].Value + htmlTdEnd); //adding first name  
                    sbmessageBody.Append(htmlTdStart + grid.Rows[i].Cells[1].Value + htmlTdEnd); //adding last name 
                    sbmessageBody.Append(htmlTdStart + grid.Rows[i].Cells[2].Value + htmlTdEnd); //adding gender
                    sbmessageBody.Append(htmlTdStart + grid.Rows[i].Cells[3].Value + htmlTdEnd); //adding salary  
                    sbmessageBody.Append(htmlTrEnd);
                }
                sbmessageBody.Append(htmlTableEnd);
                return Convert.ToString(sbmessageBody); // return HTML Table as string from this function  
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public static void Email(string htmlString)
        {

            string Excel_File = "";
            //this is to generate file
            // this is a relative path for c: drive documents folder file
            //Excel_File = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + "TestAttachmentFile" + ".xlsx";

            // this is a relative path for project bin folder file
            Excel_File= Application.StartupPath + "\\" +"TestAttachmentFile" + ".xlsx";
            
            try
            {
                MailMessage message = new MailMessage();
                SmtpClient smtp = new SmtpClient("secure.emailsrvr.com");
                message.From = new MailAddress("automation@nextservices.com","automation");
                message.To.Add("imrans@nextservices.com");
                message.Subject = "Test Send Mail Via SMTP";
                message.IsBodyHtml = true; //to make message body as html  
                //htmlString = "this is test mail";
                message.Body = htmlString +"<font><br><br><p> Regards,<br> API Team </font> ";
                //smtp.Port = 587;
                //smtp.Host = "smtp.gmail.com"; //for host  
                //smtp.Port = 465;


                //for attachement
                System.Net.Mail.Attachment attachment;
                //attachment = new System.Net.Mail.Attachment("your file path");
                attachment = new System.Net.Mail.Attachment(Excel_File);
                message.Attachments.Add(attachment);

                // host name
                string HostName = Dns.GetHostName();
                IPAddress[] ip = Dns.GetHostAddresses(HostName);
                string HostIPAddress = ip[1].ToString();

                smtp.Host = "secure.emailsrvr.com"; //for host  
                smtp.EnableSsl = true;
                smtp.UseDefaultCredentials = false;
                smtp.Credentials = new NetworkCredential("nextreport@nextservices.com", "welcome@123");
                
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.Send(message);
            }
            catch (Exception ex) {
                MessageBox.Show(ex.ToString());
            }
        }
        private void btnSendMail_Click(object sender, EventArgs e)
        {
            string htmlString = getHtml(dgvEmployee); //here you will be getting an html string  
            Email(htmlString); //Pass html string to Email function.  
        }
    }
}
