using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SharepointCRUD
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Read web title
            using(ClientContext clientContext = new ClientContext("http://vm-021-tzc02/sites/CSOMSite/"))
            {
                Web web = clientContext.Web;
                clientContext.Load(web);
                clientContext.ExecuteQuery();
                MessageBox.Show(web.Title);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Change the web title and update update the website
            using (ClientContext clientContext = new ClientContext("http://vm-021-tzc02/sites/CSOMSite/"))
            {
                Web web = clientContext.Web;
                web.Title = "My Csom Site Updated";
                web.Update();
                clientContext.ExecuteQuery();
                MessageBox.Show(web.Title);
            }
        }
        /// <summary>
        /// Method to Create List
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                // create list
                using (ClientContext clientContext = new ClientContext("http://vm-021-tzc02/sites/CSOMSite/"))
                {
                
                    Web web = clientContext.Web;
                    ListCreationInformation listCreationInformation = new ListCreationInformation();
                    listCreationInformation.Title = "EmployeeList";
                    listCreationInformation.TemplateType = (int)ListTemplateType.GenericList;
                    web.Lists.Add(listCreationInformation);
                    clientContext.ExecuteQuery();
                }
                MessageBox.Show("List Created");

            }
            catch (Exception ex)
            {
                throw;
            }

           
        }
        /// <summary>
        /// Method to Delete List 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                // Delete list 
                using (ClientContext clientContext = new ClientContext("http://vm-021-tzc02/sites/CSOMSite/"))
                {
                    Web web = clientContext.Web;
                    List list = web.Lists.GetByTitle("Employee");
                    list.DeleteObject();
                    clientContext.ExecuteQuery();
                }
                MessageBox.Show("List Deleted");
            }
            catch (Exception ex)
            {
                throw;
            }
           
        }
        /// <summary>
        /// Delete columns fron the list
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                // To delete columns from the list
                using (ClientContext clientContext = new ClientContext("http://vm-021-tzc02/sites/CSOMSite/"))
                {
                    Web web = clientContext.Web;
                    List list = web.Lists.GetByTitle("Employee");
                    Field field = list.Fields.GetByTitle("Expiry Date");
                    field.DeleteObject();
                    clientContext.ExecuteQuery();
                }
                MessageBox.Show("Column Deleted");

            }
            catch (Exception ex)
            {
                throw;
            }

        }

        /// <summary>
        /// Method to add columns to the list
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                // create column 
                using (ClientContext clientContext = new ClientContext("http://vm-021-tzc02/sites/CSOMSite/"))
                {
                    Web web = clientContext.Web;
                    List list = web.Lists.GetByTitle("Employee");
                    list.Fields.AddFieldAsXml(@"<Field 
                                           Name='Expiry Date'
                                           DisplayName='Expiry Date'
                                           Type='DateTime'
                                           Format='DateOnly'
                                           Required='FALSE'
                                           >
                                        <Default>[today]</Default>
                                      </Field>", true, AddFieldOptions.DefaultValue);
                    clientContext.ExecuteQuery();
                }
                MessageBox.Show("Column Created");
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        /// <summary>
        /// Insert items o the column
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                using (ClientContext clientContext = new ClientContext("http://vm-021-tzc02/sites/CSOMSite/"))
                {
                    Web web = clientContext.Web;
                    List myList = web.Lists.GetByTitle("EmployeeList");
                    ListItemCreationInformation listItemCreationInformation = new ListItemCreationInformation();
                    ListItem listitem = myList.AddItem(listItemCreationInformation);
                    listitem["Title"] = "FirstRow";
                    listitem["Amount"] = 100;
                    listitem["Birth_x0020_Date"] = DateTime.Now;
                    myList.Update();
                    clientContext.ExecuteQuery();
                }
                MessageBox.Show("Items Added");
            }
            catch (Exception ex)
            {
                throw;
            }

        }

        /// <summary>
        /// Add multiple columns to the list
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                DateTime start, end;
                using (ClientContext clientContext = new ClientContext("http://vm-021-tzc02/sites/CSOMSite/"))
                {
                    Random random = new Random();
                    Web web = clientContext.Web;
                    List myList = web.Lists.GetByTitle("EmployeeList");
                    ListItemCreationInformation listItemCreationInformation = new ListItemCreationInformation();
                    start = DateTime.Now;
                    for(int i = 0; i<250; i++)
                    {
                        ListItem listitem = myList.AddItem(listItemCreationInformation);
                        listitem["Title"] = Guid.NewGuid().ToString();
                        listitem["Amount"] = random.Next(1000);
                        listitem["Birth_x0020_Date"] = DateTime.Now.AddDays(random.Next(400));
                        listitem.Update();
                        
                    }
                    clientContext.ExecuteQuery();
                }
                end = DateTime.Now;
                MessageBox.Show(end.Subtract(start).TotalSeconds.ToString());
            }
            catch (Exception ex)
            {
                throw;
            }
        }

       
    }
}
