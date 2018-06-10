using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace M.JsonHelper
{
    public class JsonHelper
    {
        #region 序列化

        public static string Serialization<T>(IEnumerable<T> entity)
        {
            DataContractJsonSerializer json = new DataContractJsonSerializer(entity.GetType());
            string szJson = "";
            //序列化
            using (MemoryStream stream = new MemoryStream())
            {
                json.WriteObject(stream, entity);
                szJson = Encoding.UTF8.GetString(stream.ToArray());
            }
            return szJson;
        }

        #endregion

        #region 反序列化
        public static List<T> JSONStringToList<T>(string JsonStr)
        {
            JavaScriptSerializer Serializer = new JavaScriptSerializer();
            List<T> objs = Serializer.Deserialize<List<T>>(JsonStr);
            return objs;
        }

        public static T Deserialize<T>(string json)
        {
            T obj = Activator.CreateInstance<T>();
            using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(json)))
            {
                DataContractJsonSerializer serializer = new DataContractJsonSerializer(obj.GetType());
                return (T)serializer.ReadObject(ms);
            }
        }

        #endregion

        #region 实体对象与DataTable之间的转换
        /// <summary>
        /// 实体转换DataTable
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="entities"></param>
        /// <returns></returns>
        public static DataTable ToDataTable<T>(IEnumerable<T> entities)
        {
            try
            {
                DataTable dtReturn = new DataTable();

                PropertyInfo[] oProps = null;

                foreach (T rec in entities)
                {
                    if (oProps == null)
                    {
                        oProps = ((Type)rec.GetType()).GetProperties();

                        foreach (PropertyInfo pi in oProps)
                        {
                            Type colType = pi.PropertyType;
                            if ((colType.IsGenericType) && (colType.GetGenericTypeDefinition() == typeof(Nullable<>)))
                            {
                                colType = colType.GetGenericArguments()[0];
                            }

                            dtReturn.Columns.Add(new DataColumn(pi.Name, colType));
                        }
                    }

                    DataRow dr = dtReturn.NewRow();
                    foreach (PropertyInfo pi in oProps)
                    {
                        dr[pi.Name] = pi.GetValue(rec, null) == null ? DBNull.Value : pi.GetValue(rec, null);
                    }

                    dtReturn.Rows.Add(dr);
                }

                return dtReturn;
            }
            catch (Exception exp)
            {
                //Log.WriteException(exp);
                throw exp;
            }
        }

        /// <summary>
        /// DataTable转换为对象
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="table"></param>
        /// <returns></returns>
        public static T ToEntity<T>(DataTable table) where T : new()
        {
            T entity = new T();
            foreach (DataRow row in table.Rows)
            {
                foreach (var item in entity.GetType().GetProperties())
                {
                    if (row.Table.Columns.Contains(item.Name))
                    {
                        if (DBNull.Value != row[item.Name])
                        {
                            item.SetValue(entity, Convert.ChangeType(row[item.Name], item.PropertyType), null);
                        }
                    }
                }
            }

            return entity;
        }

        /// <summary>
        /// DataTable转换为对象数组
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="table"></param>
        /// <returns></returns>
        public static IList<T> ToEntities<T>(DataTable table) where T : new()
        {
            IList<T> entities = new List<T>();
            foreach (DataRow row in table.Rows)
            {
                T entity = new T();
                foreach (var item in entity.GetType().GetProperties())
                {
                    item.SetValue(entity, Convert.ChangeType(row[item.Name], item.PropertyType), null);
                }
                entities.Add(entity);
            }
            return entities;
        }

        #endregion
    }
}
