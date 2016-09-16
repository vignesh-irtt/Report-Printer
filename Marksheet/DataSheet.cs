using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Novacode;
namespace Marksheet
{
    public partial class DataSheet : Form
    {
        int Current;
        string testno,sem, branch, testperiod,year;
        Docx doc;
        int i;
        
        public DataSheet(String a,string b,string c,string d,string e)
        {
            InitializeComponent();
            testno = a;
            year = b;
            sem = c;
            branch = d;
            testperiod = e;
            doc = new Document();
            this.Resize += DataSheet_Resize;
            dataGridView1.ColumnHeaderMouseDoubleClick += new DataGridViewCellMouseEventHandler(column);
            dataGridView1.RowHeaderMouseDoubleClick += new DataGridViewCellMouseEventHandler(row);
            dataGridView1.Size = new Size(this.Size.Width,this.Size.Height-(this.Size.Height%100));
        }

        void printit_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            
        }

        private void DataSheet_Resize(object sender, EventArgs e)
        {
            dataGridView1.Size = new Size(this.Size.Width, this.Size.Height - (this.Size.Height/4));
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        private void column(object s, DataGridViewCellMouseEventArgs a)
        {
            
            Current = a.ColumnIndex;
            contextMenuStrip2.Show(dataGridView1, a.Location);
        }
        private void row(object s, DataGridViewCellMouseEventArgs a)
        {
            Current = a.RowIndex;
            contextMenuStrip1.Show(dataGridView1, a.Location);

        }


        private void removeRowToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void addRowToolStripMenuItem_Click(object sender, EventArgs e)
        {


        }


        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (Current == (dataGridView1.ColumnCount) - 1)
            {
                MessageBox.Show("Last Column should be always attendance !Sorry");
                return;
            }
                dataGridView1.Columns.Add(toolStripTextBox3.Text.ToString(), toolStripTextBox3.Text.ToString());
            toolStripTextBox3.Text = "";
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            try
            {
                if (Current == (dataGridView1.ColumnCount) - 1)
                {
                    MessageBox.Show("Last Column should be always attendance !Sorry");
                    return;
                }
                if (Current == 0)
                {
                    MessageBox.Show("First two columns should not be disturbed !Sorry");
                    return;
                }
                if (Current == 1)
                {
                    MessageBox.Show("First two columns should not be disturbed !Sorry");
                    return;
                }
                if (Current == 2)
                {
                    MessageBox.Show("First three columns should not be disturbed !Sorry");
                    return;
                }
                dataGridView1.Columns.RemoveAt(Current);
            }
            catch (Exception h)
            {
                MessageBox.Show(h.Message);
            }
        }

        private void addARowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Insert(Current, 1);
        }

        private void removeTheRowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.Rows.RemoveAt(Current);
               }
            catch (Exception h)
            {
                MessageBox.Show(h.Message);
            }

        }

        private void addRowsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Insert(Current,Convert.ToInt16(toolStripTextBox2.Text.ToString()));
        }

        private void insertAColumnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Current == (dataGridView1.ColumnCount) - 1)
            {
                MessageBox.Show("Last Column should be always attendance !Sorry");
                return;
            }
            if (Current == 0)
            {
                MessageBox.Show("First two columns should not be disturbed !Sorry");
                return;
            }
            if (Current ==  1)
            {
                MessageBox.Show("First two columns should not be disturbed !Sorry");
                return;
            }
            
            DataGridViewColumn newcolumn = new DataGridViewColumn();
            newcolumn.Name = toolStripTextBox3.Text.ToString();
            newcolumn.HeaderText = toolStripTextBox3.Text.ToString();
            newcolumn.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView1.Columns.Insert(Current+1, newcolumn);
        }

        private void label1_Click(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            doc = new Document();
            Cursor.Current = Cursors.WaitCursor;
            Section header=doc.AddSection();
            int j = 0;
            dataGridView1.RefreshEdit();
            foreach (DataGridViewRow student in dataGridView1.Rows)
            {
                j++;
                if (j == dataGridView1.Rows.Count)
                    continue;
                Paragraph head = header.AddParagraph();
                TextRange range = head.AppendText("\nINSTITUTE OF ROAD AND TRANSPORT TECHNOLOGY,ERODE-638316");
                range.CharacterFormat.FontName = "Times New Roman";
                range.CharacterFormat.FontSize = 14;
                head.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                head.AppendText("\n");
                range = head.AppendText("DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING");
                range.CharacterFormat.FontSize = 14;
                head.AppendText("\n");
                range = head.AppendText("INTERNAL TEST-" + testno);
                range.CharacterFormat.FontSize = 14;
                head.AppendText("\n");
                head = header.AddParagraph();
                range = head.AppendText("Name:"+student.Cells[1].Value.ToString());
                range.CharacterFormat.FontSize = 14;
                head.AppendText("\t\t\t\t\t\t\t\t\t");
                range = head.AppendText("Roll No:"+student.Cells[2].Value.ToString());
                range.CharacterFormat.FontSize = 14;
                head.AppendBreak(BreakType.LineBreak);
                head.AppendBreak(BreakType.LineBreak);
                range = head.AppendText("Class and Branch:" + sem);
                range.CharacterFormat.FontSize = 14;
                range = head.AppendText("th");
                range.CharacterFormat.SubSuperScript = Spire.Doc.Documents.SubSuperScript.SuperScript;
                range = head.AppendText(" Semester / " + branch);
                head.AppendBreak(BreakType.LineBreak);
                range.CharacterFormat.FontSize = 14;
                Table t = header.AddTable();
                t.PreferredWidth = new PreferredWidth(WidthType.Percentage, 100);
                TableRow r = t.AddRow();
                t.TableFormat.Borders.BorderType = Spire.Doc.Documents.BorderStyle.Single;
                r.IsHeader = true;
                TableCell c = r.AddCell();
                range = c.AddParagraph().AppendText("S.No");
                range.CharacterFormat.FontSize = 14;
                c.SetCellWidth(10, Spire.Doc.CellWidthType.Percentage);
                c = r.AddCell();
                range = c.AddParagraph().AppendText("Subject Name");
                range.CharacterFormat.FontSize = 14;
                c.SetCellWidth(40, Spire.Doc.CellWidthType.Percentage);
                range.OwnerParagraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                c = r.AddCell();
                range = c.AddParagraph().AppendText("Marks Scored\n(out of 100)");
                range.CharacterFormat.FontSize = 14;
                c.SetCellWidth(25, Spire.Doc.CellWidthType.Percentage);
                c = r.AddCell();
                range = c.AddParagraph().AppendText("Remarks");
                range.CharacterFormat.FontSize = 14;
                c.SetCellWidth(25, Spire.Doc.CellWidthType.Percentage);
                for (i = 3; i < student.Cells.Count - 1; i++)
                {
                    r = t.AddRow(4);
                    t.TableFormat.Borders.BorderType = Spire.Doc.Documents.BorderStyle.Single;
                    c = r.Cells[0];
                    range = c.AddParagraph().AppendText(Convert.ToString(i-2));
                    range.CharacterFormat.FontSize = 14;
                    c = r.Cells[1];
                    range = c.AddParagraph().AppendText(student.Cells[i].OwningColumn.HeaderText);
                    range.CharacterFormat.FontSize = 14;
                    range.OwnerParagraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    c = r.Cells[2];
                    range = c.AddParagraph().AppendText(student.Cells[i].Value.ToString());
                    range.CharacterFormat.FontSize = 14;
                    c = r.Cells[3];
                    if (Convert.ToInt32(student.Cells[i].Value.ToString()) >= 50)
                        range = c.AddParagraph().AppendText("Pass");
                    else
                        range = c.AddParagraph().AppendText("Fail");
                    range.CharacterFormat.FontSize = 14;
                    
                }
                head = header.AddParagraph();
                range = head.AppendText("\nPercentage of Attendance:"+student.Cells[i].Value.ToString()+"\t\t\t\t\t(period up to "+testperiod+")\n");
                range.CharacterFormat.FontSize = 14;
                head.AppendText("\n");
                range = head.AppendText("\tKindly sign and send the acknowledgement to The Principal,Institute of Road and Transport Technology.Sri Vasavi college post.Erode-638316.\n");
                range.CharacterFormat.FontSize = 14;
                head.AppendText("\n");
                range = head.AppendText("Station:Erode\n");
                range.CharacterFormat.FontSize = 14;
                head.AppendText("\n");
                range = head.AppendText("Date:" + DateTime.Today.ToShortDateString() + "\t\t\t\t\t\t\tCLASS ADVISOR");
                range.CharacterFormat.FontSize = 14;
                t = header.AddTable();
                t.ResetCells(1, 1);
                t.PreferredWidth = new PreferredWidth(WidthType.Percentage, 100);
                r = t.Rows[0];
                r.RowFormat.Borders.Right.BorderType = Spire.Doc.Documents.BorderStyle.Cleared;
                r.RowFormat.Borders.Left.BorderType = Spire.Doc.Documents.BorderStyle.Cleared;
                r.RowFormat.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Cleared;
                r.RowFormat.Borders.Bottom.BorderType = Spire.Doc.Documents.BorderStyle.Thick;
                head = header.AddParagraph();
                range = head.AppendText("Please tear along this line");
                range.CharacterFormat.FontSize = 14;
                head.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                head = header.AddParagraph();
                head.AppendBreak(BreakType.LineBreak);
                head.AppendBreak(BreakType.LineBreak);
                range = head.AppendText("Department : " + branch + " and Engineering");
                range.CharacterFormat.FontSize = 14;
                head.AppendBreak(BreakType.LineBreak);
                head.AppendBreak(BreakType.LineBreak);
                range = head.AppendText("Class : " + sem);
                range.CharacterFormat.FontSize = 14;
                range = head.AppendText("th");
                range.CharacterFormat.SubSuperScript = Spire.Doc.Documents.SubSuperScript.SuperScript;
                range = head.AppendText(" Semester/ "+year+" year");
                range.CharacterFormat.FontSize = 14;
                head.AppendBreak(BreakType.LineBreak);
                head.AppendBreak(BreakType.LineBreak);
                head = header.AddParagraph();
                range = head.AppendText("ACKNOWLEDGEMENT");
                range.CharacterFormat.FontSize = 14;
                head.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                head = header.AddParagraph();
                head.AppendBreak(BreakType.LineBreak);
                range = head.AppendText("\tI recieved the progress report for Internal " + testno + " examination dated on ____________ of my Son/Daughter _____________");
                range.CharacterFormat.FontSize = 14;
                range = head.AppendText(" Roll No ____________________ sent by you.");
                range.CharacterFormat.FontSize = 14;
                head.AppendBreak(BreakType.LineBreak);
                head.AppendBreak(BreakType.LineBreak);
                range = head.AppendText("Remarks of the parent if any:");
                range.CharacterFormat.FontSize = 14;
                head.AppendBreak(BreakType.LineBreak);
                head.AppendBreak(BreakType.LineBreak);
                range = head.AppendText("Station:");
                range.CharacterFormat.FontSize = 14;
                head.AppendBreak(BreakType.LineBreak);
                head.AppendBreak(BreakType.LineBreak);
                range = head.AppendText("Date:" + DateTime.Today.ToShortDateString() + "\t\t\t\t\t\tSignature of the parent/Guardian");
                range.CharacterFormat.FontSize = 14;
                if (j+1 != dataGridView1.Rows.Count)
                   head.AppendBreak(BreakType.PageBreak);
            }
            Cursor = Cursors.Default;
            saveFileDialog1.Filter = "docx files (*.docx)|*.docx|pdf files (*.pdf*)|*.pdf*";
            DialogResult file=saveFileDialog1.ShowDialog();
            if(file.ToString().CompareTo("OK")==0)
            if(saveFileDialog1.FilterIndex==1)
                doc.SaveToFile(saveFileDialog1.FileName.ToString(), FileFormat.Docx);
            else
                doc.SaveToFile(saveFileDialog1.FileName.ToString()+".pdf", FileFormat.PDF);
        }
                
            
        

        
        
        private void button2_Click(object sender, EventArgs e)
        {
            preview(dataGridView1.Rows[0]);
               
        }
        private void preview(DataGridViewRow student)
        {
            
        }
        private void print(DataGridViewRow student)
        {
            
        }
    }
}
