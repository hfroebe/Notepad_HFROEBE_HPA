using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Windows.Forms;

namespace Notepad_HFROEBE
{
    class Program
    {

       public static AutomationTasks AT = new AutomationTasks();

        static void Main(string[] args)
        {
            NotepadTask();
        }


        
        static void NotepadTask()
        {


            try
            {
                IntPtr handleWindow;

                Process[] proc = Process.GetProcessesByName("notepad");
                if (proc.Length == 0)
                {

                    Process.Start("Notepad.exe");
                    Thread.Sleep(200);

                    Console.WriteLine("Step 1 successfully completed. Notepad is checked to make sure it is NOT open.");
                    Console.WriteLine("------------------------ ");
                }
                else
                {
                    Console.WriteLine("Step 2: Please save your work, close Notepad and try again.");
                    Console.ReadLine();
                    return;
                }

                Console.WriteLine("Step 2 successfully completed. Notepad is opened by application. ");
                Console.WriteLine("------------------------ ");


                Process p = Process.GetProcessesByName("notepad").FirstOrDefault();
                handleWindow = p.MainWindowHandle;




                var window = AutomationElement.FromHandle(handleWindow);

                var conditionMenuBar = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuBar);


                var menuBar = window.FindAll(TreeScope.Descendants, conditionMenuBar)
                 .OfType<AutomationElement>().FirstOrDefault(ui => !ui.Current.Name.Contains("Bar"));


                var conditionFileMenu = new AndCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem),
                    new PropertyCondition(AutomationElement.NameProperty, "File")
                );

                
                var fileMenu = menuBar.FindFirst(TreeScope.Children, conditionFileMenu);


                if (fileMenu is null)
                {
                    Console.WriteLine("********Manual Interference detected. No menu bar found. Close the application and try again *********** ");
                    Console.ReadLine();
                    return;
                }



                
                bool resultClickNew = AT.InvokeSubMenuItemByName(fileMenu, "New", true);

                if (!resultClickNew)
                {
                    Console.WriteLine("********Manual Interference detected. Close the application and try again *********** ");
                    Console.ReadLine();
                    return;
                }
                else
                {
                    Console.WriteLine("Step 3 successfully completed. New is clicked. ");
                    Console.WriteLine("------------------------ ");
                }


                string windowTitle = "Untitled - Notepad";
                AutomationElement notepadRootElement = AT.getRootElement(windowTitle);

                Condition conditionNotepadRootElement = new AndCondition(
                      new PropertyCondition(AutomationElement.IsEnabledProperty, true),
                      new PropertyCondition(AutomationElement.ControlTypeProperty,
                      ControlType.Document)
                      );


                AutomationElementCollection elementCollection =
                   notepadRootElement.FindAll(TreeScope.Children, conditionNotepadRootElement);


                foreach (AutomationElement el in elementCollection)
                {
                    el.SetFocus();
                    SendKeys.SendWait("Hello World!");
                   


                }

                Console.WriteLine("Step 4 successfully completed. Text entered. ");
                Console.WriteLine("------------------------ ");

                System.IO.Directory.CreateDirectory("C:\\Task_Output"); //line ignored if directory exists


                bool resultClickSaveAs = AT.InvokeSubMenuItemByName(fileMenu, "Save As...", true);

                if (!resultClickSaveAs)
                {
                    Console.WriteLine("********Manual Interference detected. Close the application and try again *********** ");
                    Console.ReadLine();
                    return;
                }

                else
                {
                    Console.WriteLine("Step 5 successfully completed. Path entered and Save As is clicked. ");
                    Console.WriteLine("------------------------ ");
                }


                Thread.Sleep(2000);

                SendKeys.SendWait("C:\\Task_Output\\Hello_World.txt");


                var saveAsDialog = AT.GetAutomationElement(notepadRootElement, "Save As");


                var btnSave = AT.GetAutomationElement(notepadRootElement, "Save", "button");

                Console.WriteLine("please wait..");
                Console.WriteLine("------------------------ ");

                InvokePattern targetInvokePattern = btnSave.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;

                if (targetInvokePattern == null)
                {
                    Console.WriteLine("********Manual Interference detected. Close the application and try again *********** ");
                    Console.ReadLine();
                    return;
                }


                targetInvokePattern.Invoke();

               

                if (btnSave == null)
                {
                    Console.WriteLine("********Manual Interference detected. Close the application and try again *********** ");
                    Console.ReadLine();
                    return;
                }

                else
                {
                    Console.WriteLine("Step 6 successfully completed. Save is clicked. ");
                    Console.WriteLine("------------------------ ");
                }


                

                try
                {



                    
                    var confirmSaveAsDialog = AT.GetAutomationElement(notepadRootElement, "Confirm");               
                    var btnConfirmYes = AT.GetAutomationElement(notepadRootElement, "Yes", "button");

                    btnConfirmYes.SetFocus();
                    if (btnConfirmYes != null)
                    {
                        InvokePattern targetInvokePatternConfirm = btnConfirmYes.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;

                        if (targetInvokePatternConfirm == null)

                        {
                            Console.WriteLine("********Manual Interference detected. Close the application and try again *********** ");
                            Console.ReadLine();
                            return;
                        }

                        targetInvokePatternConfirm.Invoke();

                        Console.WriteLine("Step 7 successfully completed. Save confirmed. ");
                        Console.WriteLine("------------------------ ");

                    }

                    else
                    {
                        Console.WriteLine("********Manual Interference detected. Close the application and try again *********** ");
                        Console.ReadLine();
                        return;
                    }
                }
                catch (Exception e) when (e is Win32Exception || e is FileNotFoundException || e is NullReferenceException)
                {
                    
                    Console.WriteLine("The following exception was raised: ");
                    Console.WriteLine(e.Message);
                    Console.WriteLine("********Manual Interference detected. Close the application and try again *********** ");
                    Console.ReadLine();
                    return;
                }

               



                Thread.Sleep(150);
                var myProcesses = Process.GetProcessesByName("notepad");

                foreach (var process in myProcesses)
                {
                    process.CloseMainWindow();
                    process.Close();
                }


                if (File.Exists(@"C:\Task_Output\Hello_World.txt"))
                {
                    Console.WriteLine(@"Step 8 successfully completed. The file exists at C:\Task_Output\Hello_World.txt");

                    Console.WriteLine("------------------------ ");

                    Console.WriteLine("Task completed successfully. You may now Close console window.");

                    Console.ReadLine();
                }

            }
            catch (Exception e)
            {

                Console.WriteLine("The following exception was raised: ");
                Console.WriteLine(e.Message);
                Console.WriteLine("*********Manual Interference detected. Close the application and try again *********** ");
                Console.ReadLine();
                return;
            }

        }
    }
}
