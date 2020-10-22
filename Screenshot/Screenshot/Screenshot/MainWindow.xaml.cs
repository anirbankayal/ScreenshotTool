/*
 * 
 * This code will take screenshot and save it in a folder in your c:\Screenshot\... folder
 * After exiting the application it will create a zip of the folder and will also save the images in a document also
 * The window will be available in buttom right cornet of the screen
 * Developer: Anirban Kayal (anirban.kayal@tcs.com) - Emp# 397851
 * 
 */


using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Interop;

namespace Screenshot
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		static int number = 0;
		static int docNum = 0;
		static string screenshotFolder = "";
		static string globalPath = "";

		public MainWindow()
		{
            try
            {

                InitializeComponent();
                this.ResizeMode = System.Windows.ResizeMode.NoResize;
                this.WindowStyle = System.Windows.WindowStyle.ToolWindow;
                if(ValidateFramework() && ValidateOffice())
                {
                    createFolder();
                }
                else
                {
                    MainWindow mw = new MainWindow();
                    mw.Close();
                    string ErrorMessage = string.Empty;
                    if(!ValidateFramework())
                    {
                        ErrorMessage = ErrorMessage + ".NET Framework is not installed or not functioning in your system. \n";
                    }
                    if(!ValidateOffice())
                    {
                        ErrorMessage = ErrorMessage + "MS Office is not installed or corrupted in your system. \n";
                    }

                    System.Windows.Forms.MessageBox.Show(ErrorMessage);
                }
            }
            catch(Exception ex)
            {
                MainWindow mw = new MainWindow();
                mw.Close();
                string ErrorMessage = "Something went wrong... ";
                System.Windows.Forms.MessageBox.Show(ErrorMessage);
            }
			
		}


        public bool ValidateFramework()
        {
            try
            {
                string systemVersionVal = System.Runtime.InteropServices.RuntimeEnvironment.GetSystemVersion().ToString();
                if(String.IsNullOrEmpty(systemVersionVal))
                    return false;
                else
                    return true;
            }
            catch(Exception ex)
            {
                return false;
            }
        }

        public bool ValidateOffice()
        {
            try
            {
                Type officeType = Type.GetTypeFromProgID("Word.Application");
                if(officeType == null)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch(Exception ex)
            {
                return false;
            }
        }

        /// <summary>
        /// Button event after clicking capture button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
		public void capture_click(object sender, RoutedEventArgs e)
		{
			try
			{
				Bitmap bitmap = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);

				Graphics graphics = Graphics.FromImage(bitmap as Image);

				graphics.CopyFromScreen(0, 0, 0, 0, bitmap.Size);

				bitmap.Save(getname(), ImageFormat.Jpeg);
				docNum++;
			}
			catch (Exception ex)
			{
				MainWindow mw = new MainWindow();
				mw.Close();
				string ErrorMessage = "An error occoured: \n" + ex.ToString();
                System.Windows.Forms.MessageBox.Show(ErrorMessage);
			}
		}

        /// <summary>
        /// Logic for naming of generated files
        /// </summary>
        /// <returns></returns>
		public string getname()
		{
			try
			{
				screenshotFolder = globalPath;
				string name = "";
				string date = DateTime.Now.ToString();
				string folderName = globalPath;
				string filename = "\\file_" + number.ToString() + ".jpeg";
				name = folderName + filename;
				number++;
				return name;
			}

			catch (Exception ex)
			{
				string ErrorMessage = "An error occoured in creating file name: \n" + ex.ToString();
				System.Windows.Forms.MessageBox.Show("");
				return null;
			}
		}


        /// <summary>
        /// Logic for zipping the files
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
		public void createZip(object sender, CancelEventArgs e)
		{
            //Thread thread = new Thread(new ThreadStart(CreateDocument));
            //thread.Start();

            var task = new Task(CreateDocument);
            task.Start();

			if (docNum > 0)
			{
				try
				{
					string startPath = globalPath;
					string zipPath = globalPath + ".zip";

					ZipFile.CreateFromDirectory(startPath, zipPath);
					System.Windows.Forms.MessageBox.Show("Zipped in " + zipPath + "\r\n" + "Will be saved as .docx in few seconds.");
				}
				catch (Exception ex)
				{
					MainWindow mw = new MainWindow();
					mw.Close();
					string ErrorMessage = "An error occoured in creating zip file: \n" + ex.ToString();
					System.Windows.Forms.MessageBox.Show("");
				}
				finally
				{
                    Task.WaitAll(task);
				}
			}
			else
			{
				Directory.Delete(globalPath);
				System.Windows.Forms.MessageBox.Show("You have not captured any Screenshot. Closing Application.");
			}
		}

        /// <summary>
        /// Logic for creating folder
        /// </summary>
		public void createFolder()
		{
			this.Topmost = true;
			this.Top = Screen.PrimaryScreen.Bounds.Height - 90;
			this.Left = Screen.PrimaryScreen.Bounds.Width - 100;

			try
			{
				int i = 0;
				for (int j = 0; j <= 100; j++)
				{

					string ExtnPath = i.ToString();
					string path = @"C:\Screenshot\Screenshot_" + DateTime.Now.Date.ToShortDateString().Replace("/", "") + "_" + i;  

					if (Directory.Exists(path))
					{
						i++;
						continue;
					}
					else
					{
						Directory.CreateDirectory(path);
						globalPath = path;
						break;
					}
				}
			}
			catch (Exception ex)
			{
				string ErrorMessage = "An error occoured in creating folder: \n" + ex.ToString();
				System.Windows.Forms.MessageBox.Show("");
			}

		}

        /// <summary>
        /// Logic for creating document of images
        /// </summary>
		public void CreateDocument()
		{
			try
			{
				Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();
				winword.Visible = false;
				object missing = System.Reflection.Missing.Value;				
				Microsoft.Office.Interop.Word.Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);

                //Microsoft.Office.Interop.Word.Range title = winword.ActiveDocument.Range(0, 0);
                //title.Text = "Screenshot: " + DateTime.Now.ToString() + "\u2028\n";

                //document.BuiltInDocumentProperties("Title").Value = "Screenshot: " + DateTime.Now.ToString();
				//Save the document
				object filename = getDocname();
				string foldername = globalPath;

				for (int i = docNum-1; i >= 0; i--)
				{
					string imageName = "\\file_" + i.ToString() + ".jpeg";
					string file = foldername + imageName;
					document.InlineShapes.AddPicture(file, Type.Missing, Type.Missing, Type.Missing);
                    Microsoft.Office.Interop.Word.Range rng = winword.ActiveDocument.Range(0, 0);
                    rng.Text = "\u2028\n";
				}

				document.SaveAs2(ref filename);
				document.Close(ref missing, ref missing, ref missing);
				document = null;
				winword.Quit(ref missing, ref missing, ref missing);
				winword = null;
			}
			catch (Exception ex)
			{
				System.Windows.Forms.MessageBox.Show(ex.Message);
			}
		}

        /// <summary>
        /// Logic for Document name
        /// </summary>
        /// <returns></returns>
		public string getDocname()
		{
			try
			{
				screenshotFolder = @"C:\Screenshot";
				string name = "";
				int i = 0;
				

				for (int j = 0; j <= 100; j++)
				{

					string filename = "\\Doc_" + DateTime.Now.Date.ToShortDateString().Replace("/", "") + "_" + i.ToString() + ".docx";
					name = screenshotFolder + filename;

					if (File.Exists(name))
					{
						i++;
						continue;
					}
					else
					{ 
						break;
					}
				}
				
				return name;
			}

			catch (Exception ex)
			{
				string ErrorMessage = "An error occoured in creating document file name: \n" + ex.ToString();
                System.Windows.Forms.MessageBox.Show(ErrorMessage);
				return null;
			}
		}
	}
}
