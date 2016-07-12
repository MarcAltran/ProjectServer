using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    class Program
    {
        static void Main(string[] args)
        {
            ConsoleColor initialColor = Console.ForegroundColor;
            try
            {
                // First, very simple: list Projects, and their tasks
                //var projectContext = NeosSDI.ProjectOnline.CSOM.ProjectCSOMManager.ReadProjects();
                //Console.ForegroundColor = ConsoleColor.Yellow;

                //Console.WriteLine("Found projects:");
                //Console.ForegroundColor = initialColor;

                //foreach (var p in projectContext.Projects)
                //{
                //    Console.WriteLine(p.Name);
                //    Console.WriteLine("Project Tasks:");
                //    foreach (var t in p.Tasks)
                //    { 
                //        Console.WriteLine(t.Name + ": " + t.Id);
                //        Console.WriteLine(NeosSDI.ProjectOnline.CSOM.ProjectCSOMManager.getPredecessor(t.Id.ToString()));
                //    }
                //}

                // More tricky: add a task, and assign me to it, from Today
                //Console.ForegroundColor = ConsoleColor.Yellow;
                //Console.WriteLine("Creating Task and Assignment...");
                //Console.ForegroundColor = initialColor;
                //NeosSDI.ProjectOnline.CSOM.ProjectCSOMManager.AddTasksToProject();

                //########################################################################
                //### MLL: Hem de descomentar les següents línies per fer la imputació ###
                //########################################################################

                //Console.ForegroundColor = ConsoleColor.Yellow;
                //Console.WriteLine("Adding Actual On Assignment...");
                //Console.ForegroundColor = initialColor;
                //NeosSDI.ProjectOnline.CSOM.ProjectCSOMManager.AddActualToTaskTimeSheet();

                //########################################################################
                //NeosSDI.ProjectOnline.CSOM.ProjectCSOMManager.ChangeRemainingWork();


                //##############################################################################
                //### MLL: Hem de descomentar les següents línies per fer la export a Oracle ###
                //##############################################################################
                NeosSDI.ProjectOnline.CSOM.ProjectCSOMManager.ConnectOracleAndSetLot();
                NeosSDI.ProjectOnline.CSOM.ProjectCSOMManager.getProjectServerData();



            }
            catch (Exception ex)
            {
                Console.WriteLine("Error:");
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                if (ex.InnerException != null)
                    Console.WriteLine(ex.InnerException.Message);

                Console.ForegroundColor = initialColor;
            }
            Console.WriteLine("Press a key...");
            Console.ReadLine();
        }
    }
}
