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

        /// <summary>
        /// method to add discription
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                // Add List Description
                using (ClientContext clientContext = new ClientContext("http://vm-021-tzc02/sites/CSOMSite/"))
                {
                    Web web = clientContext.Web;
                    List list = web.Lists.GetByTitle("Employee");
                    list.Description = "Added Discription";
                    list.Update();
                    clientContext.ExecuteQuery();
                }
                MessageBox.Show("List Description Added");
            }
            catch (Exception ex)
            {
                throw;
            }

        }

        /// <summary>
        /// Method to update title
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                // Add List Description
                using (ClientContext clientContext = new ClientContext("http://vm-021-tzc02/sites/CSOMSite/"))
                {
                    Web web = clientContext.Web;
                    List list = web.Lists.GetByTitle("AddedTitle");
                    list.Title = "Employee";
                    list.Update();
                    clientContext.ExecuteQuery();
                }
                MessageBox.Show("List Title Added");
            }
            catch (Exception ex)
            {
                throw;
            }

        }
        /// <summary>
        /// List versioning enabled
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                // Add List Description
                using (ClientContext clientContext = new ClientContext("http://vm-021-tzc02/sites/CSOMSite/"))
                {
                    Web web = clientContext.Web;
                    List list = web.Lists.GetByTitle("Employee");
                    list.EnableVersioning = true;
                    list.Update();
                    clientContext.ExecuteQuery();
                }
                MessageBox.Show("List Versioning Enabled");
            }
            catch (Exception ex)
            {
                throw;
            }

        }

        /// <summary>
        /// Method to enable minor versioning 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button12_Click(object sender, EventArgs e)
        {
            try
            {
                // Add List Description
                using (ClientContext clientContext = new ClientContext("http://vm-021-tzc02/sites/CSOMSite/"))
                {
                    Web web = clientContext.Web;
                    List list = web.Lists.GetByTitle("Employee");
                    list.EnableMinorVersions = true;
                    list.Update();
                    clientContext.ExecuteQuery();
                }
                MessageBox.Show("List Minor Versioning Enabled");
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        /// <summary>
        /// Method to enable Major versioning
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button13_Click(object sender, EventArgs e)
        {
            try
            {
                // Add List Description
                using (ClientContext clientContext = new ClientContext("http://vm-021-tzc02/sites/CSOMSite/"))
                {
                    Web web = clientContext.Web;
                    List list = web.Lists.GetByTitle("Employee");
                    list.EnableVersioning = true;
                    list.MajorVersionLimit = 200;
                    list.Update();
                    clientContext.ExecuteQuery();
                }
                MessageBox.Show("List Major Versioning with limit Enabled");
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        /// <summary>
        /// Method to enable force checkout
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button14_Click(object sender, EventArgs e)
        {
            try
            {
                // Add List Description
                using (ClientContext clientContext = new ClientContext("http://vm-021-tzc02/sites/CSOMSite/"))
                {
                    Web web = clientContext.Web;
                    List list = web.Lists.GetByTitle("Employee");
                    list.ForceCheckout = true;
                    list.Update();
                    clientContext.ExecuteQuery();
                }
                MessageBox.Show("Force CHeckout Enabled");
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        /// <summary>
        /// Method to enable folder creation
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button15_Click(object sender, EventArgs e)
        {
            try
            {
                // Add List Description
                using (ClientContext clientContext = new ClientContext("http://vm-021-tzc02/sites/CSOMSite/"))
                {
                    Web web = clientContext.Web;
                    List list = web.Lists.GetByTitle("Employee");
                    list.EnableFolderCreation = true;
                    list.Update();
                    clientContext.ExecuteQuery();
                }
                MessageBox.Show("Folder CReation Enabled");
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        /// <summary>
        /// To enable quick launch
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button16_Click(object sender, EventArgs e)
        {
            try
            {
                // Add List Description
                using (ClientContext clientContext = new ClientContext("http://vm-021-tzc02/sites/CSOMSite/"))
                {
                    Web web = clientContext.Web;
                    List list = web.Lists.GetByTitle("Employee");
                    list.OnQuickLaunch = true;
                    list.Update();
                    clientContext.ExecuteQuery();
                }
                MessageBox.Show("Quick launch Enabled");
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        /// <summary>
        /// /TO anbale content types
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button17_Click(object sender, EventArgs e)
        {
            try
            {
                // Add List Description
                using (ClientContext clientContext = new ClientContext("http://vm-021-tzc02/sites/CSOMSite/"))
                {
                    Web web = clientContext.Web;
                    List list = web.Lists.GetByTitle("Employee");
                    list.ContentTypesEnabled = true;
                    list.Update();
                    clientContext.ExecuteQuery();
                }
                MessageBox.Show("Content types Enabled");
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            try
            {
                using (ClientContext clientContext = new ClientContext("http://vm-021-tzc02/sites/CSOMSite/"))
                {
                    Web web = clientContext.Web;
                    List list = web.Lists.GetByTitle("Employee");
                    clientContext.Load(list);
                    clientContext.ExecuteQuery();
                    ContentTypeCollection contentTypes = clientContext.Site.RootWeb.ContentTypes;
                    clientContext.Load(contentTypes);
                    clientContext.ExecuteQuery();

                    ContentType contentType = contentTypes.Where(c => c.Name == "Timecard").First();
                    list.ContentTypes.AddExistingContentType(contentType);
                    list.Update();
                    clientContext.ExecuteQuery();
                }
                MessageBox.Show("Content types Added");
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            try
            {
                using (ClientContext clientContext = new ClientContext("http://vm-021-tzc02/sites/CSOMSite/"))
                {
                    Web web = clientContext.Web;
                    List list = web.Lists.GetByTitle("Employee");
                    clientContext.Load(list);
                    clientContext.Load(list.ContentTypes);
                    clientContext.ExecuteQuery();
                    ContentTypeCollection contentTypes = list.ContentTypes;
                    clientContext.Load(contentTypes);
                    clientContext.ExecuteQuery();

                    ContentType contentType = contentTypes.Where(c => c.Name == "Timecard").First();
                    contentType.DeleteObject();
                    list.Update();
                    clientContext.ExecuteQuery();
                }
                MessageBox.Show("Content type removed from the list");
            }
            catch (Exception ex)
            {
                throw;
            }
        }

       

        private void button20_Click_1(object sender, EventArgs e)
        {
            try
            {
                using (ClientContext clientContext = new ClientContext("http://vm-021-tzc02/sites/CSOMSite/"))
                {
                    Web web = clientContext.Web;
                    List list = web.Lists.GetByTitle("Employee");
                    clientContext.Load(list);
                    clientContext.Load(list.ContentTypes);
                    clientContext.ExecuteQuery();
                    ContentTypeCollection contentTypes = list.ContentTypes;
                    clientContext.Load(contentTypes);
                    clientContext.ExecuteQuery();

                    for (int i = 0; i < contentTypes.Count; i++)
                    {
                        ContentType contentType = contentTypes[i];
                        contentType.DeleteObject();
                    }

                    list.Update();
                    clientContext.ExecuteQuery();
                }
                MessageBox.Show("All Content type removed from the list");
            }
            catch (Exception ex)
            {
                throw;
            }
        }
    }
}
