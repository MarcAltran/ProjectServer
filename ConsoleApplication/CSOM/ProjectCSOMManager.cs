using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.ProjectServer.Client;
using Microsoft.SharePoint.Client;
using System.DirectoryServices;

//using Oracle.DataAccess.Client;
using System.Data.OracleClient;

using System.Data;
using System.Data.SqlClient;
using System.Globalization;
//using NeosSDI.ProjectOnline.Business;

namespace NeosSDI.ProjectOnline.CSOM
{
    public class ProjectCSOMManager
    {
        #region Variables
        // Change this value if you are on prem and not Online
        private static int timeOut = 20;
        private static string projName = "Project"; // define the name of your tenant here

        // MLL: Variables para cargar los datos de la imputación
        private static List<string> username = new List<string>();
        private static List<string> usernameGuid = new List<string>();
        private static List<string> usernameDistinc = new List<string>();
        private static List<string> taskid = new List<string>();
        private static List<string> projectid = new List<string>();
        private static List<int> actualwork = new List<int>();
        private static List<DateTime> fecha = new List<DateTime>();

        // MLL: Variables para cargar los datos concretos del Timesheet
        private static List<string> TimesheetTaskId = new List<string>();
        private static List<DateTime> TimesheetTaskDate = new List<DateTime>();
        private static List<int> TimesheetActualwork = new List<int>();

        private static List<string> TaskUIDList = new List<string>();

        // MLL: Variables per a formatejar les dates
        private static DateTime dt = DateTime.Now;
        private static DateTime wkStDt = DateTime.MinValue;
        private static DateTime wkStDt2 = DateTime.MinValue;

        //MLL: Variables per a Oracle
        private static string connectionstring = "";
        private static int id_dom = 1;
        private static int id_lot = 0;
        private static int id_entry = 0;


        private static string projDomain = string.Format("{0}.onmicrosoft.com", projName);

        // Set the Project Server client context.
        private static ProjectContext projContext;

        #endregion

        private static string PwaPath
        {
            get
            {
                    return "http://intranet/pwa";

            }
        }

        #region Metodos que no se usan
        /// <summary>
        /// This method performs a very simple operation: Read Projects, and the Tasks of these projets
        /// Result is stored in the ProjectContext, and returned to the client.
        /// Step by step, we have to:
        /// - Manage authentification, for Project Online, or Project On Prem
        /// - Create the Query, to ask for the projects, and to include some additionnal properties (dates, tasks...)
        /// - Execute the Query
        /// </summary>
        /// <returns></returns>
        public static ProjectContext ReadProjects()
        {
            try
            {
                projContext = new ProjectContext(PwaPath);


                projContext.Credentials = new System.Net.NetworkCredential("marc.lluis", "4ltr4n@2016", "ps");

                // Use IncludeWithDefaultProperties to force CSOM to load the Tasks collection, otherwize we have a (very) lazy loading
                // Careful: the Load method does not perform the Load ! It prepare the context before the ExecuteQuery is run
                projContext.Load(projContext.Projects,
                    c => c.IncludeWithDefaultProperties(pr => pr.StartDate, pr => pr.FinishDate, pr => pr.Tasks));

                // Actual execution of the Load - AFter this method, the Projects collection contains data, and the properties which are specified below. 
                projContext.ExecuteQuery();


            }
            catch (Exception ex)
            {
                throw ex;
            }

            return projContext;
        }

        /// <summary>
        /// This method add a task to the project, and assign, me to it.
        /// The date of assignment is hardcoded to today
        /// 
        /// The steps:
        /// - Manage authentification, for Project Online, or Project On Prem
        /// - Prepare the Queries to load Projects and the Web Context (to get the current user)
        /// - Execute this first queries
        /// - Prepare the Query to load the Resource linked to the current user
        /// - Load the First existing Project, and check it out, to get its Draft version
        /// - Create the Task with the TaskCreationInformation class
        /// - Add it to the Project
        /// - Update the Project, and execute this long query
        /// - Create an assignment for this task, and the current resource, with the AssignmentCreationInformation class
        /// - Add it to the Project
        /// - Update the Project, and execute this query
        /// - Publish/Checkin the project
        /// </summary>
        public static void AddTasksToProject()
        {
            try
            {
                projContext = new ProjectContext(PwaPath);

                projContext.Credentials = new System.Net.NetworkCredential("marc.lluis", "4ltr4n@2016", "ps");

                // Use IncludeWithDefaultProperties to force CSOM to load the Tasks collection, otherwize we have a lazy loading
                // Careful: the Load method does not perform the Load ! It prepare the context before the ExecuteQuery is run.
                projContext.Load(projContext.Projects,
                    c => c.IncludeWithDefaultProperties(pr => pr.StartDate, pr => pr.FinishDate, pr => pr.Tasks));

                projContext.Load(projContext.Web.CurrentUser);
                projContext.ExecuteQuery();
                string currentUserName = projContext.Web.CurrentUser.LoginName;

                // Important to exclude resource without associated user (a resource who does not have an account)
                projContext.Load(projContext.EnterpriseResources,
                        res => res.IncludeWithDefaultProperties(r => r.User, r => r.User.LoginName).Where(r => r.User != null && r.User.LoginName == currentUserName));

                // Actual execution of the Load - After this method, the Projects collection contains data, and the properties which are specified below. 
                projContext.ExecuteQuery();

                var pubProject = projContext.Projects.FirstOrDefault();

                var currentProject = pubProject.CheckOut();

                if (currentProject == null)
                    throw new Exception("Please create a project !");

                var currentResource = projContext.EnterpriseResources.FirstOrDefault();
                if (currentResource == null)
                    throw new Exception("Please create yourself as a resource !");

                TaskCreationInformation tsk = new TaskCreationInformation();
                tsk.Name = string.Format("Task created at {0}", DateTime.Now);
                tsk.Start = DateTime.Now;
                tsk.Finish = DateTime.Now.AddDays(3);
                var newTask = currentProject.Tasks.Add(tsk);
                projContext.Load(newTask);

                QueueJob qj = currentProject.Update();
                JobState js = projContext.WaitForQueue(qj, timeOut);
                projContext.ExecuteQuery();
                AssignmentCreationInformation ass = new AssignmentCreationInformation();
                ass.ResourceId = currentResource.Id;

                ass.TaskId = newTask.Id;
                ass.Start = newTask.Start;
                ass.Finish = newTask.Finish;
                currentProject.Assignments.Add(ass);
                currentProject.Update();
                qj = currentProject.Publish(true);
                js = projContext.WaitForQueue(qj, timeOut);
                projContext.ExecuteQuery();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.Message);
                throw ex;
            }
        }
        #endregion


        /// <summary>
        /// This method add actuals to an existing assignment, by using TimeSheets.
        /// Pre requisite: periods must be created
        /// Current bug: the timehseet must exist for this period (simply clic on TimeSheet link in PWA once for the current period
        /// Steps are:
        /// - Manage authentification, for Project Online, or Project On Prem
        /// - Prepare the Query to load the Web Context (to get the current user)
        /// - Execute this first query
        /// - Prepare the Query to load the Resource linked to the current user
        /// - Prepare the Query to load the current period, by including, TimeSheet, which includes Lines, which include Work and Assignments. The lines are filtered in order to include only Standard Lines, and not admin lines (sick, vacation...)
        /// - Execute this Query
        /// - For the current line, retrieve the planned work for today 
        /// - Create the TimesheetWork with the TimeSheetWorkCreationInformation class, and set the different properties
        /// - Add this work to the line, and Update the TimeSheet
        /// - Submit the TimeSheet, after management of the current status of the TimeSheet
        /// </summary>
        public static void AddActualToTaskTimeSheet()
        {


            try
            {
                //projContext = new ProjectContext(PwaPath);
                //if (IsProjectOnline)
                //    projContext.ExecutingWebRequest += ClaimsHelper.clientContext_ExecutingWebRequest;
                //else
                // MLL: A mi me llega un USER_ID de la base de datos --> Montar un metodo que busque el UserID y devuelva las credenciales
                // MLL: Hay que pasar el USUARIO a imputar en las tarea mediante parametro!!


                //MLL: Esta variable la puedo rellenar de la COLUMNA == ResourceNTAccount --> de la TABLA == [MSP_EpmResource]
                //string currentUserName = projContext.Web.CurrentUser.LoginName;
                String domain = System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties().DomainName;
                domain = domain.Substring(0, domain.IndexOf('.'));

                //Obtenemos los datos de la imputación mediante estas dos funciones
                getData();
                getTaskID();

                //Hacemos un bucle para poder loguearnos por cada usuario, en el caso de que fuera distinto
                for (int a = 0; a < usernameDistinc.Count; a++)
                {
                    string currentUserName = usernameDistinc.ElementAt(a);
                    string currentUserName2 = usernameDistinc.ElementAt(a).Substring(8 + domain.Length);
                    projContext = new ProjectContext(PwaPath);


                    //projContext.Credentials = new System.Net.NetworkCredential(currentUserName2, "4ltr4n@2016", domain);
                    projContext.Credentials = new System.Net.NetworkCredential("marc.lluis", "4ltr4n@2016", "ps");


                    projContext.Load(projContext.Web.CurrentUser);
                    projContext.ExecuteQuery();

                    // Important to exclude resource without associated user (a resource who does not have an account)
                    projContext.Load(projContext.EnterpriseResources,
                            res => res.IncludeWithDefaultProperties(r => r.User, r => r.User.LoginName).Where(r => r.User != null && r.User.LoginName == currentUserName));

                    // Actual execution of the Load - After this method, the Projects collection contains data, and the properties which are specified below. 
                    projContext.ExecuteQuery();


                    var currentResource = projContext.EnterpriseResources.FirstOrDefault();
                    if (currentResource == null)
                    {
                        AsignarUsuarioATarea(projectid.ElementAt(a),usernameGuid.ElementAt(a),taskid.ElementAt(a),fecha.ElementAt(a), fecha.ElementAt(a));
                        throw new Exception("Please create yourself as a resource !");
                    }

                    projContext.Load(projContext.TimeSheetPeriods, c => c.Where(p => p.Start <= DateTime.Now && p.End >= DateTime.Now).
                        IncludeWithDefaultProperties(p => p.TimeSheet,
                                                     p => p.TimeSheet.Lines.Where(l => l.LineClass == TimeSheetLineClass.StandardLine).
                        IncludeWithDefaultProperties(l => l.Assignment,
                                                           l => l.Assignment.Task,
                                                           l => l.Work)));

                    projContext.ExecuteQuery();

                    var myPeriod = projContext.TimeSheetPeriods.FirstOrDefault();

                    if (myPeriod == null)
                        throw new Exception("Please create the periods in your server settings");

                    

                    //var line = myPeriod.TimeSheet.Lines;// FirstOrDefault();


                    foreach (var l in myPeriod.TimeSheet.Lines)
                    {

                        // Hay que pasar el GUID de la tarea mediante parametro!!
                        for (int i = 0; i < TimesheetTaskId.Count; i++)
                        {
                            if (l.Id == Guid.Parse(TimesheetTaskId.ElementAt(i)))
                            {
                                var plannedwork = l.Work.Where(w => w.Id == Guid.Parse(TimesheetTaskId.ElementAt(i))).FirstOrDefault();

                                TimeSheetWorkCreationInformation workCreation = new TimeSheetWorkCreationInformation
                                {

                                    ActualWork = string.Format("{0}h", TimesheetActualwork.ElementAt(i)), // Horas trabajadas
                                    Start = TimesheetTaskDate.ElementAt(i),
                                    End = TimesheetTaskDate.ElementAt(i),
                                    Comment = "From CSOM",
                                    NonBillableOvertimeWork = "0",
                                    NonBillableWork = "0",
                                    OvertimeWork = "0",
                                    PlannedWork = plannedwork == null ? "0h" : plannedwork.PlannedWork
                                };

                                l.Work.Add(workCreation);

                                myPeriod.TimeSheet.Update();
                            }

                        }
                    }

                    if (myPeriod.TimeSheet.Status == TimeSheetStatus.Approved ||
                        myPeriod.TimeSheet.Status == TimeSheetStatus.Submitted ||
                        myPeriod.TimeSheet.Status == TimeSheetStatus.Rejected)
                    {
                        myPeriod.TimeSheet.Recall();
                    }

                    myPeriod.TimeSheet.Submit("GO");

                    projContext.ExecuteQuery();
                }
            } //Fin del bucle de autenticación por cada usuario

            catch (Exception ex)
            {

                System.Diagnostics.Trace.WriteLine(ex.Message);
                throw ex;
            }
        }

        /*public static void ChangeRemainingWork()
        {
            getData();

            // Get the list of projects on the server.
            projContext.Load(projContext.Projects);
            projContext.ExecuteQuery();

            //MLL: Cambiar la obtención del nombre del proyecto por el GUID que nos vendra de la BDD
            var proj = projContext.Projects.First(p => p.Name == "Project");
            projContext.ExecuteQuery();

            var draftProj = proj.CheckOut();

            projContext.Load(draftProj.Tasks);
            projContext.ExecuteQuery();

            //CreateNewTask(draftProj);


            foreach (DraftTask task in draftProj.Tasks)
            {

                Console.WriteLine("\n\t GUID: {0} \n\t TASK NAME: {1} \n\t DURATION: {2} \n\t ACTUAL WORK: {3} \n\t REMAINING WORK: {4}", task.Id.ToString(), task.Name, task.Duration, task.ActualWork, task.RemainingWork);

                Guid mainGuid = Guid.Parse(taskid);

                if (task.Id == mainGuid)
                {
                    task.RemainingDuration = "300h";
                    draftProj.Update();

                }

            }

            draftProj.Publish(true);
            QueueJob qJob = projContext.Projects.Update();
            JobState jobState = projContext.WaitForQueue(qJob, 200);
        }*/

        #region Helper Methods
        /// <summary>
        /// Funcionalitats que ajuden a obtenir valors de la BDD i de crear la connexió.
        /// </summary>

        private static string GetFullUserName(string userName)
        {
            return string.Format("i:0#.f|membership|{0}@{1}", userName, projDomain);
        }

        public static string getPredecessor(string taskid)
        {
            String sentencia = "SELECT LINK_PRED_UID FROM [ProjectWebApp].[pub].[MSP_LINKS] WHERE LINK_SUCC_UID = '" + taskid + "'";
            SqlDataReader rs = GestioBBDD.ExecutarConsulta(sentencia);
            if (rs.Read())
            {
                return rs.GetValue(0).ToString();
            }
            else
            {
                return "";
            }
        }

        public static string getSucessor(string taskid)
        {
            String sentencia = "SELECT LINK_SUCC_UID FROM [ProjectWebApp].[pub].[MSP_LINKS] WHERE LINK_SUCC_UID = '" + taskid + "'";
            SqlDataReader rs = GestioBBDD.ExecutarConsulta(sentencia);
            if (rs.Read())
            {
                return rs.GetValue(0).ToString();
            }
            else
            {
                return "";
            }
        }

        public static void getData()
        {
            String sentencia = "SELECT TASK_ID, USER_ID, TASK_DATE, TASK_REMAINING_WORK, TASK_ACTUAL_WORK FROM MSP_Redmine_Tasks";
            SqlDataReader rs = GestioBBDD.ExecutarConsulta(sentencia);
            if (rs.HasRows)
            {
                while (rs.Read())
                {
                    taskid.Add(rs.GetValue(0).ToString());
                    usernameGuid.Add(rs.GetValue(1).ToString());

                    sentencia = "SELECT ResourceNTAccount FROM MSP_EpmResource WHERE ResourceUID ='" + rs.GetValue(1).ToString() + "'";
                    SqlDataReader rs2 = GestioBBDD.ExecutarConsulta(sentencia);
                    if (rs2.Read())
                        username.Add(rs2.GetValue(0).ToString());

                    sentencia = "SELECT ProjectUID FROM MSP_EpmTask WHERE TaskUID = '" + rs.GetValue(0).ToString() + "'";
                    SqlDataReader rs3 = GestioBBDD.ExecutarConsulta(sentencia);
                    if (rs3.Read())
                        projectid.Add(rs3.GetValue(0).ToString());

                    fecha.Add(Convert.ToDateTime(rs.GetValue(2).ToString()));
                    //Console.WriteLine("TASK REMAINING WORK: " + rs.GetValue(3).ToString());
                    actualwork.Add(Convert.ToInt32(rs.GetValue(4).ToString()));
                }
            }

            usernameDistinc = username.Distinct().ToList();
        }

        public static void getTaskID()
        {

            wkStDt = dt.AddDays(1 - Convert.ToDouble(dt.DayOfWeek));
            wkStDt2 = wkStDt.AddDays(6);
            string diaInicioSemana = wkStDt.Date.ToString("yyyy-MM-dd");
            string diaFinSemana = wkStDt2.Date.ToString("yyyy-MM-dd");


            //Inicializamos la lista para que no se repitan los registros al llamar la función 2 veces
            TimesheetTaskId.Clear();

            for (int i = 0; i < taskid.Count; i++)
            {
                String sentencia = "SELECT Tl.[TimesheetLineUID]"
                            + " FROM dbo.MSP_EpmTask task"
                            + " INNER JOIN[MSP_TimesheetTask] Times ON task.TaskUID = Times.TaskUID"
                            + " INNER JOIN[MSP_TimesheetLine] Tl ON Times.[TaskNameUID] = Tl.[TaskNameUID]"
                            + " WHERE TASK.TaskUID = '" + taskid.ElementAt(i) + "'"
                            + " AND CAST('" + fecha.ElementAt(i).Date.ToString("yyyy-MM-dd") + "' AS DATE) BETWEEN CAST('" + diaInicioSemana + "' AS DATE) AND CAST('" + diaFinSemana + "' AS DATE)"
                            + " AND CAST(Tl.CreatedDate AS DATE) BETWEEN CAST('" + diaInicioSemana + "' AS DATE) AND CAST('" + diaFinSemana + "' AS DATE)";
                //and CAST(Tl.CreatedDate AS DATE) = CAST('" + fechadesdesemana + "' AS DATE)"
                SqlDataReader rs = GestioBBDD.ExecutarConsulta(sentencia);
                if (rs.HasRows)
                {
                    while (rs.Read())
                    {
                        TimesheetTaskId.Add(rs.GetValue(0).ToString());
                        TimesheetTaskDate.Add(fecha.ElementAt(i));
                        TimesheetActualwork.Add(actualwork.ElementAt(i));
                    }
                }
            }

        }

        public static void getProjectServerData()
        {
            DateTime fechainicio = DateTime.MinValue;
            DateTime fechafinal = DateTime.MinValue;

            String select = "select issue_ext, exported from MSP_EPMProjectServer_Tasks";
            SqlDataReader rs1 = GestioBBDD.ExecutarConsulta(select);

            if (rs1.HasRows)
            {
                while (rs1.Read())
                {
                    if (Convert.ToBoolean(rs1.GetValue(1)))
                    {
                        TaskUIDList.Add(rs1.GetValue(0).ToString());
                    }
                }
            }

            String create = " select * into saved from MSP_EPMProjectServer_Tasks where Exported = 1";
            SqlDataReader rs = GestioBBDD.ExecutarConsulta(create);

            String truncate = "truncate table MSP_EPMProjectServer_Tasks";
            rs = GestioBBDD.ExecutarConsulta(truncate);

            String insertSaved = "insert into MSP_EPMProjectServer_Tasks"
                                + " select * from saved";
            rs = GestioBBDD.ExecutarConsulta(insertSaved);

            String DeleteTemp = "DROP TABLE dbo.saved";
            rs = GestioBBDD.ExecutarConsulta(DeleteTemp);


            String sentencia = "SELECT"
                                + " t.TaskUID AS 'ID',"
                                + " t.TaskName AS 'NOM',"
                                + " t.TaskDuration AS 'HORAS ESTIMADES',"
                                + " t.TaskStartDate AS 'DATA INICI',"
                                + " t.TaskFinishDate AS 'DATA FI',"
                                + " t.TaskPercentCompleted AS 'PERCENTATGE COMPLETAT',"
                                + " t.TaskPriority AS 'PRIORITAT DE LA TASCA',"
                                + " r.resourcename AS 'NOM RECURS',"
                                + " t.TaskParentUID AS 'ID TASCA PARE',"
                                + " p.[Codi aplicacio] AS 'Codi Aplicacio',"
                                + " p.[Job ABS] AS 'Job ABS'"
                            + " FROM MSP_EpmTask t"
                            + " FULL OUTER JOIN MSP_EPMASSIGNMENT ea ON t.taskuid = ea.taskuid"
                            + " FULL OUTER JOIN MSP_EPMRESOURCE r ON ea.resourceuid = r.resourceuid"
                            + " INNER JOIN MSP_EPMPROJECT ep ON ep.ProjectUID = t.ProjectUID"
                            + " INNER JOIN MSP_EpmProject_UserView p ON p.ProjectUID = ep.ProjectUID"
                            + " WHERE T.TaskWBS is not NULL";

            rs = GestioBBDD.ExecutarConsulta(sentencia);
            if (rs.HasRows)
            {
                while (rs.Read())
                {
                    fechainicio = DateTime.Parse(rs.GetValue(3).ToString());
                    fechafinal = DateTime.Parse(rs.GetValue(4).ToString());

                    if (rs1.HasRows && TaskUIDList.Count == 0)
                    {
                        for (int i = 0; i < TaskUIDList.Count; i++)
                        {
                            if (!TaskUIDList.ElementAt(i).ToLower().Equals(rs.GetValue(0).ToString()))
                            {
                                try
                                {
                                    string predecesor = getPredecessor(rs.GetValue(0).ToString());
                                    string sucesor = getSucessor(rs.GetValue(0).ToString());

                                    String insert = "INSERT INTO[dbo].[MSP_EPMProjectServer_Tasks] ([issue_ext], [project_ext], [subject], [description], [estimated_hours], [start_date]"
                                                        + ", [due_date], [done_ratio], [priority_ext], [assigned_to_ext], [parent_issue_ext], [status_ext], [tracker_ext], [job_abs], [Exported], predecesor, sucesor)"
                                                    + "VALUES ("
                                                    + "'" + rs.GetValue(0) + "', '" + rs.GetValue(9) + "', "
                                                    + "'" + rs.GetValue(1) + "', '', " + rs.GetValue(2).ToString().Replace(",", ".") + ", '" + fechainicio.ToString("yyyy-MM-dd") + "', '" + fechafinal.ToString("yyyy-MM-dd") + "', " + rs.GetValue(5) + ", "
                                                    + "'" + rs.GetValue(6) + "', '" + rs.GetValue(7) + "', '" + rs.GetValue(8) + "', NULL, NULL, '" + rs.GetValue(10) + "', 0,'" + predecesor + "', '"+ sucesor + "')";
                                    SqlDataReader rs2 = GestioBBDD.ExecutarConsulta(insert);

                                    if (InsertDataOracle(rs.GetValue(0).ToString(), rs.GetValue(9).ToString(), rs.GetValue(1).ToString(), Convert.ToInt32(rs.GetValue(2)), fechainicio, fechafinal,
                                                    Convert.ToInt32(rs.GetValue(5)), rs.GetValue(6).ToString(), rs.GetValue(7).ToString(), rs.GetValue(8).ToString(), rs.GetValue(10).ToString(), predecesor))
                                    {
                                        String update = "UPDATE MSP_EPMProjectServer_Tasks"
                                                        + "SET Exported = 1"
                                                        + "WHERE issue_ext = '" + rs.GetValue(0) + "'";

                                    }
                                }
                                catch (Exception ex)
                                {

                                    System.Diagnostics.Trace.WriteLine(ex.Message);
                                    throw ex;
                                }
                            }
                        }
                    }

                    else
                    {
                        try
                        {
                            string predecesor = getPredecessor(rs.GetValue(0).ToString());
                            string sucesor = getSucessor(rs.GetValue(0).ToString());

                            String insert = "INSERT INTO[dbo].[MSP_EPMProjectServer_Tasks] ([issue_ext], [project_ext], [subject], [description], [estimated_hours], [start_date]"
                                                + ", [due_date], [done_ratio], [priority_ext], [assigned_to_ext], [parent_issue_ext], [status_ext], [tracker_ext], [job_abs], [Exported], predecesor, sucesor)"
                                            + "VALUES ("
                                            + "'" + rs.GetValue(0) + "', '" + rs.GetValue(9) + "', "
                                            + "'" + rs.GetValue(1) + "', '', " + rs.GetValue(2).ToString().Replace(",", ".") + ", '" + fechainicio.ToString("yyyy-MM-dd") + "', '" + fechafinal.ToString("yyyy-MM-dd") + "', " + rs.GetValue(5) + ", "
                                            + "'" + rs.GetValue(6) + "', '" + rs.GetValue(7) + "', '" + rs.GetValue(8) + "', NULL, NULL, '" + rs.GetValue(10) + "', 0,'" + predecesor + "', '" + sucesor + "')";
                            SqlDataReader rs2 = GestioBBDD.ExecutarConsulta(insert);

                            if (InsertDataOracle(rs.GetValue(0).ToString(), rs.GetValue(9).ToString(), rs.GetValue(1).ToString(), Convert.ToInt32(rs.GetValue(2)), fechainicio, fechafinal,
                                            Convert.ToInt32(rs.GetValue(5)), rs.GetValue(6).ToString(), rs.GetValue(7).ToString(), rs.GetValue(8).ToString(), rs.GetValue(10).ToString(), predecesor))
                            {
                                String update = "UPDATE MSP_EPMProjectServer_Tasks"
                                                + "SET Exported = 1"
                                                + "WHERE issue_ext = '" + rs.GetValue(0) + "'";

                            }
                        }
                        catch (Exception ex)
                        {

                            System.Diagnostics.Trace.WriteLine(ex.Message);
                            throw ex;
                        }
                    }
                }

                Console.WriteLine("Importacion de datos de Project Server ==> Lista!");
            }
        }

        private static void AsignarUsuarioATarea(string proj_id, string resource, string task_id, DateTime start, DateTime finish)
        {
            var proj = projContext.Projects.First(p => p.Id == Guid.Parse(proj_id));
            projContext.ExecuteQuery();

            var draftProj = proj.CheckOut();

            AssignmentCreationInformation ass = new AssignmentCreationInformation();
            ass.ResourceId = Guid.Parse(resource);

            ass.TaskId = Guid.Parse(task_id);
            ass.Start = start;
            ass.Finish = finish;
            draftProj.Assignments.Add(ass);
            draftProj.Update();

            projContext.Load(draftProj.Tasks);

            projContext.ExecuteQuery();
        }

        public static class GestioBBDD
        {
            public static SqlDataReader ExecutarConsulta(string sql)
            {
                SqlDataReader ret = null;
                SqlCommand command = null;
                try
                {
                    command = crearConexio(sql);
                    ret = command.ExecuteReader(CommandBehavior.CloseConnection);

                }
                catch (Exception e)
                {
                    if (command != null) command.Connection.Close();
                    throw e;
                }
                return ret;
            }

            private static SqlCommand crearConexio(String sql)
            {
                string connStr = System.Configuration.ConfigurationManager.ConnectionStrings["Connection"].ConnectionString;
                SqlConnection conn = new SqlConnection(connStr);
                conn.Open();
                SqlCommand command = conn.CreateCommand();
                command.CommandText = sql;
                return command;
            }


        }

        #endregion

        #region Oracle Functions

        public static void ConnectOracleAndSetLot()
        {

            connectionstring = String.Format(
                                        "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.6.144.181)" +
                                        "(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=DESorcl1)));User Id=USER_INT_TSO;Password=t2WR3.250;");



            using (OracleConnection conn = new OracleConnection(connectionstring)) // connect to oracle
            {
                OracleCommand objCmd = new OracleCommand();
                objCmd.Connection = conn;
                objCmd.CommandText = "pkg_in.ins_lot";
                objCmd.CommandType = CommandType.StoredProcedure;
                objCmd.Parameters.Add("id_dom", OracleType.Number).Value = id_dom;
                objCmd.Parameters.Add("id_lot", OracleType.Number).Direction = ParameterDirection.ReturnValue;

                try
                {
                    conn.Open();
                    objCmd.ExecuteNonQuery();
                    System.Console.WriteLine("Id Lot: {0}", objCmd.Parameters["id_lot"].Value);

                    id_lot = Convert.ToInt32(objCmd.Parameters["id_lot"].Value);

                }
                catch (Exception ex)
                {
                    conn.Close();
                    System.Console.WriteLine("Exception: {0}", ex.ToString());
                }

                conn.Close();
            }

        }

        public static bool InsertDataOracle(string issue_ext, string project_ext, string subject, int estimated_hours, DateTime start_date,
                                            DateTime due_date, int done_ratio, string priority_ext, string assigned_to_ext, string parent_issue_ext, string job_abs,
                                            string predecesor)
        {
            try
            {
                connectionstring = String.Format(
                                       "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.6.144.181)" +
                                       "(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=DESorcl1)));User Id=USER_INT_TSO;Password=t2WR3.250;");

                using (OracleConnection conn = new OracleConnection(connectionstring)) // connect to oracle
                {
                    OracleCommand objCmd2 = new OracleCommand();
                    objCmd2.Connection = conn;
                    objCmd2.CommandText = "pkg_in.ins_entry_issue";
                    objCmd2.CommandType = CommandType.StoredProcedure;

                    //objCmd2.Parameters.Add("id_dom", OracleType.Number).Value = 1;//id_dom;
                    //objCmd2.Parameters.Add("id_lot", OracleType.Number).Value = 46;//id_lot;
                    objCmd2.Parameters.Add("id_dom", OracleType.Number).Value = id_dom;
                    objCmd2.Parameters.Add("id_lot", OracleType.Number).Value = id_lot;
                    objCmd2.Parameters.Add("issue_ext", OracleType.VarChar).Value = issue_ext; //GUID de Project Server
                    objCmd2.Parameters.Add("project_ext", OracleType.VarChar).Value = project_ext; //Codi de dialeg
                    objCmd2.Parameters.Add("subject", OracleType.VarChar).Value = subject; //Task Name
                    objCmd2.Parameters.Add("description", OracleType.VarChar).Value = ""; //Task Description
                    objCmd2.Parameters.Add("estimated_hours", OracleType.Number).Value = estimated_hours; //Horas estimadas
                    objCmd2.Parameters.Add("start_date", OracleType.DateTime).Value = start_date; //Fecha de inicio
                    objCmd2.Parameters.Add("due_date", OracleType.DateTime).Value = due_date; //Fecha fin 
                    objCmd2.Parameters.Add("done_ratio", OracleType.Number).Value = done_ratio; //Porcentaje completado
                    objCmd2.Parameters.Add("priority_ext", OracleType.VarChar).Value = priority_ext; //Prioridad que viene de PS -- 500 es normal
                    objCmd2.Parameters.Add("assigned_to_ext", OracleType.VarChar).Value = assigned_to_ext; //Recurso
                    objCmd2.Parameters.Add("parent_issue_ext", OracleType.VarChar).Value = parent_issue_ext; //GUID de la tarea padre
                    objCmd2.Parameters.Add("status_ext", OracleType.VarChar).Value = ""; //Campos para otros ERP que no es project server -- Hay que informarlos a NULL / Blanco
                    objCmd2.Parameters.Add("tracker_ext", OracleType.VarChar).Value = ""; //Campos para otros ERP que no es project server -- Hay que informarlos a NULL / Blanco

                    objCmd2.Parameters.Add("id_entry", OracleType.Number).Direction = ParameterDirection.ReturnValue;
                    conn.Open();
                    objCmd2.ExecuteNonQuery();

                    id_entry = Convert.ToInt32(objCmd2.Parameters["id_entry"].Value);
                    Console.WriteLine("Id Entry: {0}", id_entry);

                    //OracleCommand objCmd3 = new OracleCommand();
                    //objCmd3.Connection = conn;
                    //objCmd3.CommandText = "pkg_in.ins_entry_issue_cv";
                    //objCmd3.CommandType = CommandType.StoredProcedure;

                    //objCmd3.Parameters.Add("id_entry", OracleType.Number).Value = id_entry;
                    //objCmd3.Parameters.Add("id_cf", OracleType.Number).Value = 51;
                    //objCmd3.Parameters.Add("valor", OracleType.VarChar).Value = job_abs;

                    //objCmd3.ExecuteNonQuery();

                    if (!predecesor.Equals(""))
                    {
                        OracleCommand objCmd4 = new OracleCommand();
                        objCmd4.Connection = conn;
                        objCmd4.CommandText = "pkg_in.ins_entry_issue_rel";
                        objCmd4.CommandType = CommandType.StoredProcedure;

                        objCmd4.Parameters.Add("id_entry", OracleType.Number).Value = id_entry;
                        objCmd4.Parameters.Add("issue_to_ext", OracleType.VarChar).Value = predecesor;
                        objCmd4.Parameters.Add("relation_type_ext", OracleType.Number).Value = 51; // 1 == FS
                        objCmd4.Parameters.Add("retard", OracleType.VarChar).Value = "";

                        objCmd4.ExecuteNonQuery();
                    }

                    //pkg_in.ins_entry_issue_rel(id_entry, 'EA_tasca_2', 1, NULL);

                    conn.Close();
                }

                return true;
            }
            catch (Exception ex)
            {
                System.Console.WriteLine("Exception: {0}", ex.ToString());
                return false;
            }
        }
        #endregion
    }
}

