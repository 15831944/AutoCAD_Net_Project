using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data.SqlClient;

namespace KFWH_CMD
{
    /// <summary>
    /// 用于数据库的增删改查
    /// </summary>
    public class SqlHelper
    {
        public static readonly string constr = @"server=10.19.80.15\CNKFEWHEALIC01;DataBase=db_MTOALT1;uid=CP;pwd=Cp654321";
        /// <summary>
        /// 增加，删除，更改数据库
        /// </summary>
        /// <param name="sqlcmd"> sql语句</param>
        /// <param name="param">sql语句使用到的参数，防止sql注入</param>
        /// <returns>受影响的行数</returns>
        public static int MyExecuteNonQuery(string sqlcmd, params SqlParameter[] param)
        {
            using (SqlConnection con = new SqlConnection(constr))
            {
                using (SqlCommand cmd = new SqlCommand(sqlcmd, con))
                {
                    con.Open();
                    if (param != null)
                    {
                        cmd.Parameters.AddRange(param);
                    }
                    return cmd.ExecuteNonQuery();
                }
            }
        }
        /// <summary>
        ///返回受影响的首行的首列值
        /// </summary>
        /// <param name="sqlcmd"> sql语句</param>
        /// <param name="param">sql语句使用到的参数，防止sql注入</param>
        /// <returns>受影响的首行的列值</returns>
        public static object MyExecuteScalar(string sqlcmd, params SqlParameter[] param)
        {
            using (SqlConnection con = new SqlConnection(constr))
            {
                using (SqlCommand cmd = new SqlCommand(sqlcmd, con))
                {
                    con.Open();
                    if (param != null)
                    {
                        cmd.Parameters.AddRange(param);
                    }
                    return cmd.ExecuteScalar(); ;
                }
            }
        }
        /// <summary>
        /// 读取数据库的内容，返回sqldatareader对象
        /// </summary>
        /// <param name="sqlcmd">sql查询语句</param>
        /// <param name="param">sql查询语句用到的参数，防止sql注入</param>
        /// <returns></returns>
        public static SqlDataReader MyExecuteReader(string sqlcmd, params SqlParameter[] param)
        {
            SqlConnection con = new SqlConnection(constr);
            using (SqlCommand cmd = new SqlCommand(sqlcmd, con))
            {
                if (param != null)
                {
                    cmd.Parameters.AddRange(param);
                }
                try
                {
                    con.Open();
                    return cmd.ExecuteReader(System.Data.CommandBehavior.CloseConnection);
                }
                catch (Exception)
                {
                    con.Close();
                    con.Dispose();
                    throw;
                }
            }
        }
    }
}
