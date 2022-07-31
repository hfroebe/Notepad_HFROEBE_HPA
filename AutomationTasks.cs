using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Automation;

namespace Notepad_HFROEBE
{
    public  class AutomationTasks
    {

         public AutomationElement getRootElement(string windowTitle)
        {
            AutomationElement root = AutomationElement.RootElement;
            AutomationElement result = null;
            foreach (AutomationElement window in root.FindAll(TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)))
            {
                try
                {
                    if (window.Current.Name.Contains(windowTitle) && window.Current.IsKeyboardFocusable)
                    {
                        result = window;
                        break;
                    }
                    else if (window.Current.AutomationId.Contains(windowTitle) && window.Current.IsKeyboardFocusable)
                    {
                        result = window;
                        break;
                    }
                    else if (window.Current.Name.Contains(windowTitle))
                    {
                        result = window;
                        break;
                    }
                }
                catch (Exception e)
                {
                    //this will throw if a window has been closed since the start of the program.. no need to stop
                }
            }

            return result;
        }


        public AutomationElement GetAutomationElement(AutomationElement parentElement, string childElementName)
        {
            try
            {
                AutomationElement chidElement = parentElement.FindAll(TreeScope.Descendants, Automation.RawViewCondition).
                                    OfType<AutomationElement>().FirstOrDefault(elm => elm.Current.Name.Contains(childElementName));
                return chidElement;
            }
            catch (Exception e)
            {

                Console.WriteLine("The following exception was raised: ");
                Console.WriteLine(e.Message);
                Console.WriteLine("********Manual Interference detected. Close the application and try again *********** ");
                Console.ReadLine();
                return null;
            }

           
            
        }

        public AutomationElement GetAutomationElement(AutomationElement parentElement, string childElementName, string childElementControlType)
        {

            try
            {
                AutomationElement chidElement = parentElement.FindAll(TreeScope.Descendants, Automation.RawViewCondition).
                    OfType<AutomationElement>().FirstOrDefault(elm => elm.Current.LocalizedControlType.Contains(childElementControlType) && elm.Current.Name.Contains(childElementName));
                return chidElement;
            }
            catch (Exception e)
            {

                Console.WriteLine("The following exception was raised: ");
                Console.WriteLine(e.Message);
                Console.WriteLine("********Manual Interference detected. Close the application and try again *********** ");
                Console.ReadLine();
                return null;
            }

            
        }


        public bool InvokeSubMenuItemByName(AutomationElement menuItem, string menuName, bool exactMatch)
        {
            var subMenus = GetMenuSubMenuList(menuItem);
            AutomationElement namedMenu = null;
            if (exactMatch)
            {
                namedMenu = subMenus.FirstOrDefault(elm => elm.Current.Name.Equals(menuName));
            }
            else
            {
                namedMenu = subMenus.FirstOrDefault(elm => elm.Current.Name.Contains(menuName));
            }
            if (namedMenu is null) return false;

            InvokeMenu(namedMenu);

            return true;
        }

        public void ExpandMenu(AutomationElement menu)
        {
            if (menu.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out object pattern))
            {
                (pattern as ExpandCollapsePattern).Expand();
            }
        }


        public List<AutomationElement> GetMenuSubMenuList(AutomationElement menu)
        {
            if (menu.Current.ControlType != ControlType.MenuItem) return null;
            ExpandMenu(menu);
            var submenus = new List<AutomationElement>();
            submenus.AddRange(menu.FindAll(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem))
                                                       .OfType<AutomationElement>().ToArray());
            return submenus;
        }

        public void InvokeMenu(AutomationElement menu)
        {
            if (menu.TryGetCurrentPattern(InvokePattern.Pattern, out object pattern))
            {
                (pattern as InvokePattern).Invoke();
            }
        }

    }
}
