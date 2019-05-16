using KFWH_CMD;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KFWH_CP_SD
{
    public static class MySQLHelper
    {
        public static List<string> GetAllProjectName()
        {
            List<string> listProjects = new List<string>();
            string sqlcommand = $"select * from tb_Schedule where Project<>''";
            SqlDataReader sdr = SqlHelper.MyExecuteReader(sqlcommand);//调用自己编写的sqlhelper类
            if (sdr.HasRows)//判断行不为空
            {
                while (sdr.Read())//循环读取数据，知道无数据为止
                {
                    if (!listProjects.Contains(sdr["Project"].ToString())) listProjects.Add(sdr["Project"].ToString());
                }
            }
            return listProjects;
        }

        internal static object GetCurProjectBlkPanels(string text)
        {
            throw new NotImplementedException();
        }

        public static List<string> GetCurProjectBlks(string projectName)
        {
            List<string> listblks = new List<string>();
            string sqlcommand = $"select * from tb_Schedule where Project='{projectName}'";
            SqlDataReader sdr = SqlHelper.MyExecuteReader(sqlcommand);//调用自己编写的sqlhelper类
            if (sdr.HasRows)//判断行不为空
            {
                while (sdr.Read())//循环读取数据，知道无数据为止
                {
                    if (!listblks.Contains(sdr["BLOCK"].ToString())) listblks.Add(sdr["BLOCK"].ToString());
                }
            }
            return listblks;
        }

        public static List<string> GetCurProjectBlkPanels(string projectName, string blkName)
        {
            List<string> listPanels = new List<string>();
            string sqlcommand = $"select * from tb_Schedule where Project='{projectName}' and BLOCK='{blkName}'";
            SqlDataReader sdr = SqlHelper.MyExecuteReader(sqlcommand);//调用自己编写的sqlhelper类
            if (sdr.HasRows)//判断行不为空
            {
                while (sdr.Read())//循环读取数据，知道无数据为止
                {
                    if (!listPanels.Contains(sdr["Panel"].ToString())) listPanels.Add(sdr["Panel"].ToString());
                }
            }
            return listPanels;
        }

        public static string GetProjectLe(string projectName)
        {
            string leName = string.Empty;
            string sqlcommand = $"select * from ProjectLE where project='{projectName}'";
            SqlDataReader sdr = SqlHelper.MyExecuteReader(sqlcommand);//调用自己编写的sqlhelper类
            if (sdr.HasRows)//判断行不为空
            {
                while (sdr.Read())//循环读取数据，知道无数据为止
                {
                    leName = sdr["LE"].ToString();
                    break;
                }
            }
            return leName;
        }

        public static string GetBlocksLe(string projectName, string blkName)
        {
            string blkleName = string.Empty;
            List<string> listPanels = new List<string>();
            string sqlcommand = $"select * from TB_blockLE where project='{projectName}' and blockname='{blkName}'";
            SqlDataReader sdr = SqlHelper.MyExecuteReader(sqlcommand);//调用自己编写的sqlhelper类
            if (sdr.HasRows)//判断行不为空
            {
                while (sdr.Read())//循环读取数据，知道无数据为止
                {
                    blkleName = sdr["blockLE"].ToString();
                    break;
                }
            }
            return blkleName;
        }

        public static string GetBlockItemNo(string projectName, string blkName)
        {
            string ItemNo = string.Empty;
            List<string> listPanels = new List<string>();
            string sqlcommand = $"select * from tb_MTO where Project='{projectName}' and BLK='{blkName}'";
            SqlDataReader sdr = SqlHelper.MyExecuteReader(sqlcommand);//调用自己编写的sqlhelper类
            if (sdr.HasRows)//判断行不为空
            {
                while (sdr.Read())//循环读取数据，知道无数据为止
                {
                    ItemNo = sdr["ItemNo"].ToString();
                    break;
                }
            }
            return ItemNo;
        }

        public static string GetUserShortName(string spmName)
        {
            string userName = string.Empty;
            string sqlcommand = $"select * from UserInfor where SPMAccount='{spmName}'";
            SqlDataReader sdr = SqlHelper.MyExecuteReader(sqlcommand);//调用自己编写的sqlhelper类
            if (sdr.HasRows)//判断行不为空
            {
                while (sdr.Read())//循环读取数据，知道无数据为止
                {
                    userName = sdr["ShortName"].ToString();
                    break;
                }
            }
            return userName;
        }

        public static string GetKFLEName(string projectName)
        {
            string kfle = string.Empty;
            string sqlcommand = $"select * from tb_SDTitleBlock where ProjectName='{projectName}'";
            SqlDataReader sdr = SqlHelper.MyExecuteReader(sqlcommand);//调用自己编写的sqlhelper类
            if (sdr.HasRows)//判断行不为空
            {
                while (sdr.Read())//循环读取数据，知道无数据为止
                {
                    kfle = sdr["KFLE"].ToString();
                    break;
                }
            }
            return kfle;
        }

        public static string GetCurProjectNumber_TitleBlock(string projectName)
        {
            string HullNumber_titleBlock = projectName;
            string sqlcommand = $"select * from tb_SDTitleBlock where ProjectName='{projectName}'";
            SqlDataReader sdr = SqlHelper.MyExecuteReader(sqlcommand);//调用自己编写的sqlhelper类
            if (sdr.HasRows)//判断行不为空
            {
                while (sdr.Read())//循环读取数据，知道无数据为止
                {
                    HullNumber_titleBlock = sdr["Others"].ToString();
                    break;
                }
            }
            return HullNumber_titleBlock;
        }

        public static PanelInfor GetPanelTitleBlockInfo(string projectName, string blockName, string paenlName)
        {
            PanelInfor curPanel = new PanelInfor() { BlockName = blockName, HullNO = MySQLHelper.GetCurProjectNumber_TitleBlock(projectName), PanelName = paenlName };
            string sqlcommand = $"select * from tb_Schedule where Project='{projectName}' and BLOCK='{blockName}' and Panel='{paenlName}'";
            SqlDataReader sdr = SqlHelper.MyExecuteReader(sqlcommand);//调用自己编写的sqlhelper类
            if (sdr.HasRows)//判断行不为空
            {
                while (sdr.Read())//循环读取数据，知道无数据为止
                {
                    curPanel.PanelDecription = sdr["DESCRIPTION"].ToString();
                    curPanel.DoneBy = MySQLHelper.GetUserShortName(sdr["DoneBy"].ToString());
                    curPanel.MSSref = sdr["MSSREFNO"].ToString().Replace(projectName + "-", "");
                    break;
                }
            }
            curPanel.ItemNo = MySQLHelper.GetBlockItemNo(projectName, blockName);
            curPanel.CheckBy = MySQLHelper.GetUserShortName(MySQLHelper.GetBlocksLe(projectName, blockName)) +
                "/" + MySQLHelper.GetUserShortName(MySQLHelper.GetProjectLe(projectName)) +
                "/" + MySQLHelper.GetKFLEName(projectName);
            return curPanel;
        }
    }
}
