/*************************************************************************************
     * CLR版本：       4.0.30319.17929
     * 类 名 称：       DateTableToList
     * 机器名称：       2016-20170206KH
     * 命名空间：       SwitchPdf
     * 文 件 名：       DateTableToList
     * 创建时间：       2017/6/2 14:11:44
     * 作    者：       LIU FANG
     * 说   明：。。。。。
     * 修改时间：
     * 修 改 人：
    *************************************************************************************/


using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace SwitchPdf
{
    public class DateTableToList
    {
        /// <summary>
        /// DataTableToList<Dictionary<string, object>> 转换为list<字典>
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static List<Dictionary<string, object>> DataTableToList(DataTable dt)
        {
            List<Dictionary<string, object>> list = new List<Dictionary<string, object>>();
            Dictionary<string, object> dic = new Dictionary<string, object>();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                foreach (DataColumn dc in dt.Columns)
                {
                    DataRow row = dt.Rows[i];
                    dic.Add(dc.ColumnName, row[dc.ColumnName]);
                }
                list.Add(dic);
                dic = new Dictionary<string, object>();
            }
            return list;
        }

        /// <summary>
        /// 获得集合实体
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static List<T> EntityList<T>(DataTable dt)
        {
            if (dt == null || dt.Rows.Count == 0)
            {
                return null;
            }
            List<T> list = new List<T>();
            T entity = default(T);
            foreach (DataRow dr in dt.Rows)
            {
                entity = Activator.CreateInstance<T>();
                PropertyInfo[] pis = entity.GetType().GetProperties();
                foreach (PropertyInfo pi in pis)
                {
                    if (dt.Columns.Contains(pi.Name))
                    {
                        if (!pi.CanWrite)
                        {
                            continue;
                        }
                        if (dr[pi.Name] != DBNull.Value)
                        {
                            Type t = pi.PropertyType;
                            if (t.FullName == "System.Guid")
                            {
                                pi.SetValue(entity, Guid.Parse(dr[pi.Name].ToString()), null);
                            }
                            else
                            {
                                pi.SetValue(entity, dr[pi.Name], null);
                            }

                        }
                    }
                }
                list.Add(entity);
            }
            return list;
        }
    }
}
