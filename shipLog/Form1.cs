using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using System.IO;

namespace shipLog
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        
        //setting up some variables we'll need later in multiple properties and/or functions
        DataTable myDt = new DataTable();
        XLWorkbook currentWb;
        IXLWorksheet currentWsh;
        string filename;

        /*
         * 
         * 
         * FUNCTIONAL METHODS
         * 
         * 
         */



        //method to return the row in an excel file that a certain item occurs at in the index
        public int atIndex(IXLWorksheet sheet, string con)
        {
            
             int item; 
             for(item = 1; item <= sheet.LastRowUsed().RowNumber(); item++)
             {
                if(sheet.Cell(item, 1).Value.ToString() == con)
                {
                    return item;
                }
             }
            
            return 0;
        }

        //method to import an excel file to a data table. 
        //I will use this later in ways that are not particularly efficient, but I couldn't think of any better ideas. Please don't judge me, I'm working on fixing it
        //UPDATE: I fixed it, but I'm leaving this message here just to brighten your data little :)
        public static DataTable ImportExcelToDT(IXLWorksheet workSheet)
        {
            DataTable dt = new DataTable();

            bool firstRow = true;
            foreach (IXLRow row in workSheet.Rows())
            {
                if (firstRow)
                {
                    foreach (IXLCell cell in row.Cells())
                    {
                        dt.Columns.Add(cell.Value.ToString());
                    }
                    firstRow = false;
                }
                else
                {
                    dt.Rows.Add();
                    int i = 0;
                    foreach (IXLCell cell in row.Cells(row.FirstCellUsed().Address.ColumnNumber, row.LastCellUsed().Address.ColumnNumber))
                    {
                        dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                        i++;
                    }
                }
            }
            return dt;
        }

        //refreshes the data grid view
        public static void xlDataGridRefresh(DataGridView dgv, XLWorkbook work, DataTable dt)
        {
            dgv.DataSource = null;
            dgv.Rows.Clear();
            dgv.DataSource = dt;
        }

        /*
         * 
         * 
         * functional objects(i.e. buttons, etc that do stuff)
         * 
         * 
         */

        //receiveing button
        private void button1_Click(object sender, EventArgs e)
        {
            
            //make sure the workbook file is selected
            if (currentWb == null)
            {
                MessageBox.Show("You must select a file!");
                return;
            }

            //make sure the ticket isn't left blank. This check isn't perfect, but a better one is in the works ;)
            if (textBox1.Text == "")
            {
                MessageBox.Show("You can't receive nothing!");
                return;
            }

            //check if the ticket has already been received or not
            if (atIndex(currentWsh, textBox1.Text) > 0)
            {
                MessageBox.Show("This ticket has already been received!");
            }
            else
            {
                //create a new row, add the ticket, and update the received date for said ticket.
                int rowNumber = currentWsh.LastRowUsed().RowNumber() + 1;
                currentWsh.Cell(rowNumber, 1).Value = textBox1.Text;
                currentWsh.Cell(rowNumber, 2).Value = DateTime.Now;
                myDt.Rows.Add(textBox1.Text, DateTime.Now);
                currentWb.Save();

                xlDataGridRefresh(dataGridView1, currentWb, myDt);

                MessageBox.Show("Ticket received successfully!");
            }

            
        }

        //shipping button
        private void button2_Click(object sender, EventArgs e)
        {
            if (currentWb == null)
            {
                MessageBox.Show("You must select a file!");
                return;
            }

            if (textBox1.Text == "")
            {
                MessageBox.Show("You can't receive nothing!");
                return;
            }

            //check if the ticket has been marked as received
            int index = atIndex(currentWsh, textBox1.Text);
            if(index > 0)
            {
                //check if the ticket has been marked as shipped already
                if(currentWsh.Cell(index, 3).Value.ToString() != "")
                {
                    MessageBox.Show("This ticket has already been shipped!");
                }
                else {
                    currentWsh.Cell(index, 3).Value = DateTime.Now;
                    myDt.Rows[index - 2]["shipped"] = DateTime.Now;
                    currentWb.Save();

                    xlDataGridRefresh(dataGridView1, currentWb, myDt);

                    MessageBox.Show("Ticket has been market as shipped!");
                }
            } else
            {
                MessageBox.Show("Can't mark ticket as shipped. Has the ticket been received?");
            }
        }

        //check ticket status button
        private void button3_Click(object sender, EventArgs e)
        {
            //make sure file is selected..
            if (currentWb == null)
            {
                MessageBox.Show("You must select a file!");
                return;
            }

            //make sure textbox isn't empty...
            if (textBox1.Text == "")
            {
                MessageBox.Show("You can't receive nothing!");
                return;
            }

            //make sure the ticket exists in the index column...
            if (atIndex(currentWsh, textBox1.Text) != 0)
            {
                //set up some variables that will be useful for readability sake
                var recieved = currentWsh.Cell(atIndex(currentWsh, textBox1.Text), 2).Value;
                var shipped = currentWsh.Cell(atIndex(currentWsh, textBox1.Text), 3).Value;


                //check if the ticket has been marked as received
                if ((recieved.ToString() != "") && (recieved != null))
                {
                    //given it's received, check if it was shipped...
                    if ((shipped.ToString() != "") && (shipped != null))
                    {
                        MessageBox.Show(string.Format("Ticket #{0} was received at {1} and shipped at {2}", textBox1.Text, recieved.ToString(), shipped.ToString()));
                    }
                    else
                    {
                        MessageBox.Show(string.Format("Ticket #0 was received at {1} but was not shipped!", textBox1.Text, recieved.ToString()));
                    }
                }
            }
            else
            {
                MessageBox.Show("Ticket does not exist!");
            }
        }

        


        // TOOLSTRIP ITEMS

        //open an existing file
        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //initialize our important variables
                filename = openFileDialog1.FileName;
                currentWb = new XLWorkbook(filename);
                currentWsh = currentWb.Worksheet(1);
                myDt = ImportExcelToDT(currentWsh);

                //UI stuff to make the program display right
                label2.Text = filename;
                dataGridView1.DataSource = myDt;
            }
        }

        //create a new file
        private void newFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK) { 
                //initialize our needed variables
                currentWb = new XLWorkbook();
                filename = saveFileDialog1.FileName;
                currentWsh = currentWb.AddWorksheet("Sheet1");

                //set up the worksheet to our specifications
                currentWsh.Cell(1, 1).Value = "ticket";
                currentWsh.Cell(1, 2).Value = "received";
                currentWsh.Cell(1, 3).Value = "shipped";

                //save everything and finish initializing our variables
                currentWb.SaveAs(filename);
                myDt = ImportExcelToDT(currentWsh);

                //update UI stuff so the program displays right
                dataGridView1.DataSource = myDt;
                label2.Text = filename;
            }
        }

        //this one is self-explanatory
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Created by Jacob Townsend \nfirst built on: 10-17-2020");
        }
        
        /*
         * 
         * 
         * stuff I didn't really use but had to keep here because if I didn't the program wouldn't run lol
         * 
         * 
         */

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
        private void label2_Click(object sender, EventArgs e)
        {

        }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void fileToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
