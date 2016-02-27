using G.Excel.Common;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;

namespace G.Excel.Logic
{
    /// <summary>
    /// Excel数据分析、统计类
    /// </summary>
    public class AnalysisLib
    {
        private volatile static AnalysisLib _instance = null;
        private static readonly object lockFlag = new object();
        private IExcel _excel;

        #region 事件
        public delegate void SetButtonEnableEventHandler(bool value);
        public event SetButtonEnableEventHandler SetButtonEnableEvent;
        public delegate void ShowMessageEventHandler(string msg);
        public event ShowMessageEventHandler ShowMessageEvent;
        #endregion

        private AnalysisLib() { }

        public static AnalysisLib CreateInstance()
        {
            if (_instance == null)
            {
                lock (lockFlag)
                {
                    if (_instance == null)
                        _instance = new AnalysisLib();
                }
            }
            return _instance;
        }

        /// <summary>
        /// 从数据源中获取班级信息
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        private List<string> GetClass(DataTable table)
        {
            List<string> classes = new List<string>();
            foreach (DataRow drow in table.Rows)
            {
                if (!classes.Contains(drow["班名称"].ToString()))
                {
                    classes.Add(drow["班名称"].ToString());
                }
            }
            return classes;
        }

        /// <summary>
        /// 从数据源中获取学生基本信息
        /// </summary>
        /// <param name="table"></param>
        /// <param name="className"></param>
        /// <returns></returns>
        private List<Student> GetStudent(DataTable table, string className)
        {
            List<Student> student = new List<Student>();
            foreach (DataRow drow in table.Rows)
            {
                if (student.Where(x => x.StudentNo.Equals(drow["学号"].ToString())).Count() == 0
                    && className.Equals(drow["班名称"].ToString()))
                {
                    student.Add(new Student()
                    {
                        SeqNo = drow["序号"].ToString(),
                        StudentNo = drow["学号"].ToString(),
                        StudentName = drow["姓名"].ToString()
                    });
                }
            }
            return student;
        }

        /// <summary>
        /// 从数据源中获取专业基本信息
        /// </summary>
        /// <param name="table"></param>
        /// <param name="className"></param>
        /// <returns></returns>
        private List<Course> GetCourse(DataTable table, string className)
        {
            List<Course> course = new List<Course>();
            foreach (DataRow drow in table.Rows)
            {
                if (course.Where(x => x.CourseID.Equals(drow["课程ID"].ToString())).Count() == 0
                    && className.Equals(drow["班名称"].ToString()))
                {
                    course.Add(new Course()
                    {
                        CourseID = drow["课程ID"].ToString(),
                        CourseName = drow["课程名称"].ToString(),
                        PaperNo = drow["试卷号"].ToString(),
                        PaperMemo = drow["试卷号备注"].ToString(),
                        Confirmed = drow["是否确认"].ToString()
                    });
                }
            }
            return course;
        }

        /// <summary>
        /// 从数据源中获取学生报考的专业基本信息
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        private List<StudentCourse> GetStudentCourse(DataTable table)
        {
            List<StudentCourse> studentCourse = new List<StudentCourse>();
            foreach (DataRow drow in table.Rows)
            {
                studentCourse.Add(new StudentCourse()
                {
                    StudentNo = drow["学号"].ToString(),
                    CourseID = drow["课程ID"].ToString(),
                });
            }
            return studentCourse;
        }

        /// <summary>
        /// 分析、统计学生的报考情况
        /// </summary>
        /// <param name="fileName"></param>
        public void AnalysisData(string fileName)
        {
            ThreadPool.QueueUserWorkItem((object o) =>
            {
                try
                {
                    if (this.SetButtonEnableEvent != null)
                    {
                        this.SetButtonEnableEvent(false);
                    }
                    if (this.ShowMessageEvent != null)
                    {
                        this.ShowMessageEvent("正在获取数据...");
                    }
                    switch (Path.GetExtension(fileName).ToLower())
                    {
                        case ".xlsx":
                            _excel = EPPlusHelper.CreateInstance();
                            break;
                        case ".xls":
                            _excel = NpoiHelper.CreateInstance();
                            break;
                    }
                    DataTable sourceTable = _excel.GetSourceFromExcel(fileName, 1);
                    List<string> classes = GetClass(sourceTable);
                    List<StudentCourse> studentCourse = GetStudentCourse(sourceTable);
                    foreach (string className in classes)
                    {
                        if (this.ShowMessageEvent != null)
                        {
                            this.ShowMessageEvent(string.Format("正在统计[ {0}班 ]数据...", className));
                        }
                        StartAnalysising(sourceTable, className, studentCourse);
                        Thread.Sleep(200);
                    }

                    #region 清除缓存数据
                    sourceTable.Clear();
                    classes.Clear();
                    studentCourse.Clear();
                    _excel.ClearWorkSheets();
                    #endregion

                    if (this.ShowMessageEvent != null)
                    {
                        this.ShowMessageEvent("数据统计完成");
                    }
                }
                catch (Exception ex)
                {
                    if (this.ShowMessageEvent != null)
                    {
                        this.ShowMessageEvent(ex.Message + "===>" + ex.StackTrace);
                    }
                }
                finally
                {
                    if (this.SetButtonEnableEvent != null)
                    {
                        this.SetButtonEnableEvent(true);
                    }
                }
            });
        }

        /// <summary>
        /// 开始分析、统计学生的报考情况
        /// </summary>
        /// <param name="sourceTable"></param>
        /// <param name="className"></param>
        /// <param name="studentCourse"></param>
        private void StartAnalysising(DataTable sourceTable, string className, List<StudentCourse> studentCourse)
        {
            #region 从数据源中获取基础数据
            List<Student> student = GetStudent(sourceTable, className);
            List<Course> course = GetCourse(sourceTable, className);
            #endregion

            DataTable table = new DataTable();
            DataRow newRow = null;
            string flag = "√";

            #region 填充数据表基础数据
            for (int i = 0; i <= course.Count + 3; i++)
            {
                table.Columns.Add();
            }

            int index = table.Columns.Count - 1;
            newRow = table.NewRow();
            newRow[0] = "序号";
            newRow[1] = "学号";
            newRow[2] = "课程ID";
            newRow[index] = "总计";
            IEnumerable<string> couseID = course.Select(x => x.CourseID);
            for (int i = 3; i < index; i++)
            {
                newRow[i] = couseID.ElementAt(i - 3);
            }
            table.Rows.Add(newRow);

            newRow = table.NewRow();
            newRow[2] = "考试号";
            IEnumerable<string> paperNo = course.Select(x => x.PaperNo);
            for (int i = 3; i < index; i++)
            {
                newRow[i] = paperNo.ElementAt(i - 3);
            }
            table.Rows.Add(newRow);

            newRow = table.NewRow();
            newRow[2] = "课程名";
            IEnumerable<string> courseName = course.Select(x => x.CourseName);
            for (int i = 3; i < index; i++)
            {
                newRow[i] = courseName.ElementAt(i - 3);
            }
            table.Rows.Add(newRow);
            #endregion

            DataRow row = table.Rows[0];
            int n = 0;
            int seqno = 1;
            #region 按学生报考的情况总计
            foreach (Student s in student)
            {
                newRow = table.NewRow();
                newRow[0] = seqno++;
                newRow[1] = s.StudentNo;
                newRow[2] = s.StudentName;

                n = 0;
                for (int i = 3; i < table.Columns.Count; i++)
                {
                    foreach (StudentCourse sc in studentCourse.Where(x => x.StudentNo == s.StudentNo))
                    {
                        if (sc.CourseID.Equals(row[i].ToString()))
                        {
                            newRow[i] = flag;
                            n++;
                        }
                    }
                }
                newRow[index] = n;
                table.Rows.Add(newRow);
            }
            #endregion

            #region 按专业总计
            newRow = table.NewRow();
            newRow[0] = "总计";
            for (int i = 3; i < index; i++)
            {
                n = 0;
                foreach (DataRow drow in table.Rows)
                {
                    if (drow[i].ToString().Equals(flag))
                    {
                        n++;
                    }
                }
                newRow[i] = n;
            }

            n = 0;
            int count = 0;
            foreach (DataRow drow in table.Rows)
            {
                int.TryParse(drow[index].ToString(), out count);
                n += count;
            }
            newRow[index] = n;

            table.Rows.Add(newRow);
            #endregion

            _excel.GenerateExcel(Directory.GetCurrentDirectory(), className, table);

            #region 清除缓存数据
            student.Clear();
            course.Clear();
            #endregion
        }
    }
}
