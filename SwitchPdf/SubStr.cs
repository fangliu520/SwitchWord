/*************************************************************************************
     * CLR版本：       4.0.30319.17929
     * 类 名 称：       SubStr
     * 机器名称：       2016-20170206KH
     * 命名空间：       SwitchPdf
     * 文 件 名：       SubStr
     * 创建时间：       2017/6/2 13:50:32
     * 作    者：       LIU FANG
     * 说   明：。。。。。
     * 修改时间：
     * 修 改 人：
    *************************************************************************************/


using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SwitchPdf
{
    public class SubStr
    {
        /// <summary>
        /// 提取数据
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
        public ApplyInfo GetInfo(string c)
        {
            ApplyInfo info = new ApplyInfo();
            string content = c;
            //姓名
            int s = c.IndexOf("客户照片") + 4;
            int e = c.IndexOf("发证机关所") - s;
            string val = "";
            if (e > 0)
            {
                val = content.Substring(s, e);
                val = val.Replace("2.", "").Replace("4.", "").Replace("\r", "").Replace("\a", "").Replace(":", "").Replace("：", "");
                string sfz = val;
                e = val.IndexOf("男") > 0 ? val.IndexOf("男") : val.IndexOf("女");
                info.xingming = val.Substring(0, e).Trim();
                //身份证
                info.shenfenzheng = sfz.Substring(e + 1, sfz.Length - e - 1).Trim();

                //如果身份证为空
                if (string.IsNullOrEmpty(info.shenfenzheng))
                {
                    Regex regex = new Regex(@"\d+", RegexOptions.ECMAScript);
                    string xm = info.xingming.Trim();
                    Match match = regex.Match(info.xingming);
                    info.shenfenzheng = match.Value;
                    if (info.shenfenzheng.Length < 18)
                    {
                        info.shenfenzheng += "X";
                    }
                    info.xingming = GetChineseWord(xm);
                }
            }
            else
            {
                s = c.IndexOf("借款人照片") + 5;
                e = (c.IndexOf("男") > 0 ? c.IndexOf("男") : c.IndexOf("女")) - s;
                val = content.Substring(s, e);
                val = val.Replace("\r", "").Replace("\a", "").Replace(":", "").Replace("：", "");
                info.xingming = val;

                s = c.IndexOf("借款人基本信息") + 7;
                e = c.IndexOf("公安局") - s;
                if (e > 0)
                {
                    val = content.Substring(s, e);
                    // val = val.Replace("\r", "").Replace("\a", "").Replace(":", "").Replace("：", "");
                    val = GetNumbers(val);
                    info.shenfenzheng = val;
                    if (info.shenfenzheng.Length < 18)
                    {
                        info.shenfenzheng += "X";
                    }
                }
            }
            //申请类型
            s = c.IndexOf("申请表") + 3;
            info.leixing = content.Substring(s, 4).Replace("(", "").Replace("（", "").Replace("\r", "").Replace("\a", "").Trim();

            if (info.leixing == "消费贷"||info.leixing=="商品贷")
            {
                //申请商品
                s = c.IndexOf("商品类型") + 4;
                e = c.IndexOf("品牌") - s;
                info.shangping = content.Substring(s, e);
                info.shangping =
                info.shangping.Replace("(1)", "").Replace("7.", "").Replace("14.", "").Replace("\r", "").Replace("\a", "").Replace(":", "").Replace("：", "").Trim();
            }
            else
            {
                info.shangping = "现金";
            }

            //申请日期
            s = c.IndexOf("申请日") + 4;
            e = c.IndexOf("For HCI internal") - s;
            info.shenqingriqi = content.Substring(s, e).Replace("\r", "").Replace("\a", "").Replace(" ", "").Replace("\t", "");
            info.shenqingriqi = info.shenqingriqi.Substring(2, 8);

            //单位名称
            s = c.IndexOf("个体全称") + 4;
            e = c.IndexOf("任职部门") - s;
            info.danwei = content.Substring(s, e).Replace("46.", "").Replace("55.", "").Replace("57.", "").Replace("\r", "").Replace("\a", "").Replace(":", "").Replace("：", "").Trim();

            if (info.danwei.Length > 50)
            {
                e = c.IndexOf("职业") - s;
                info.danwei = content.Substring(s, e).Replace("46.", "").Replace("55.", "").Replace("57.", "").Replace("\r", "").Replace("\a", "").Replace(":", "").Replace("：", "").Trim();
            }


            //单位电话
            s = c.IndexOf("办公/个体电话") + 7;
            e = c.IndexOf("分机") - s;
            info.danweidianhua = content.Substring(s, e).Replace("\r", "").Replace("\a", "").Replace(":", "").Replace("：", "").Trim();

            if (info.danweidianhua.Length > 20)
            {
                s = c.IndexOf("办公/个体电") + 7;
                e = c.IndexOf("分机") - s;
                info.danweidianhua = content.Substring(s, e).Replace("\r", "").Replace("\a", "").Replace(":", "").Replace("：", "").Trim();
            }


            //个人电话
            s = c.IndexOf("手机 :") + 4;
            e = c.IndexOf("办公/个体电话") - s;
            if (e > 0)
            {
                info.shoujihao = content.Substring(s, e);
                info.shoujihao = info.shoujihao.Replace("75.", "").Replace("62.", "").Replace("\r", "").Replace("\a", "").Trim();
            }

            if (string.IsNullOrEmpty(info.shoujihao) || info.shoujihao.Length > 20)
            {
                s = c.IndexOf("电子邮箱") + 4;
                e = c.IndexOf("婚姻状况") - s;
                info.shoujihao = content.Substring(s, e).Replace("10.", "").Replace("62.", "").Replace("\r", "").Replace("\a", "").Replace(":", "").Replace("：", "").Trim();
                if (info.shoujihao.Length > 11)
                {
                    info.shoujihao = info.shoujihao.Substring(0, 11);
                }
            }

            //亲属姓名
            s = c.IndexOf("家庭成员姓名") + 6;
            e = c.IndexOf("家庭成员类型") - s;
            string qinStr = "";
            if (e > 0)
            {
                info.qinshuxingming = content.Substring(s, e).Replace("72.", "").Replace("85.", "").Replace("\r", "").Replace("\a", "").Replace(":", "").Replace("：", "").Trim();
            }
            else
            {
                s = c.IndexOf("家庭成员电话") + 6;
                e = c.LastIndexOf("收入及其它") - s;
                if (e > 0)
                {
                    val = content.Substring(s, e).Replace("67.", "").Replace("VIII.", "").Replace("\r", "").Replace("\a", "").Replace(":", "").Replace("：", "").Trim();
                    qinStr = val;
                    info.qinshuxingming = GetChineseWord(val).Replace("父亲", "").Replace("母亲", "").Replace("兄弟", "").Replace("姐妹", "").Replace("亲戚", "").Replace("同事", "").Replace("配偶", "");

                }
                if (string.IsNullOrEmpty(info.qinshuxingming))
                {
                    s = c.IndexOf("联系人姓名") + 6;
                    e = c.LastIndexOf("与申请人关系") - s;
                    info.qinshuxingming = content.Substring(s, e).Replace("67.", "").Replace("76.", "").Replace("\r", "").Replace("\a", "").Replace(":", "").Replace("：", "").Trim();
                }
            }
            if (!string.IsNullOrEmpty(qinStr))
            {
                info.qinshushouji = GetNumbers(qinStr);
                info.guanxi = qinStr.Replace(info.qinshuxingming, "").Replace(info.qinshushouji, "");
            }

            if (string.IsNullOrEmpty(info.guanxi))
            {
                //与本人关系
                s = c.IndexOf("家庭成员类型") + 6;
                e = c.IndexOf("家庭成员电话") - s;
                info.guanxi = content.Substring(s, e).Replace("73.", "").Replace("86.", "").Replace("\r", "").Replace("\a", "").Replace(":", "").Replace("：", "").Trim();
                if (info.guanxi.Length > 20)
                {
                    s = c.LastIndexOf("与申请人关系") + 6;
                    e = c.LastIndexOf("联系电话") - s;
                    info.guanxi = content.Substring(s, e).Replace("68.", "").Replace("77.", "").Replace("\r", "").Replace("\a", "").Replace(":", "").Replace("：", "").Trim();
                }
            }
            if (string.IsNullOrEmpty(info.qinshushouji))
            {
                //亲属手机
                s = c.IndexOf("家庭成员电话") + 6;
                e = c.IndexOf("家庭成员联系地址") - s;
                if (e > 0)
                {
                    info.qinshushouji = content.Substring(s, e).Replace("74.", "").Replace("87.", "").Replace("\r", "").Replace("\a", "").Replace(":", "").Replace("：", "").Trim();
                }
                else
                {
                    s = c.LastIndexOf("联系电话") + 4;
                    e = c.LastIndexOf("提供的文档") - s;
                    info.qinshushouji = content.Substring(s, e).Replace("74.", "").Replace("X.", "").Replace("\r", "").Replace("\a", "").Replace(":", "").Replace("：", "").Trim();
                }
            }
            return info;

        }
        /// <summary>
        /// 正则提取中文
        /// </summary>
        /// <param name="oriText"></param>
        /// <returns></returns>
        public string GetChineseWord(string oriText)
        {
            string x = @"[\u4E00-\u9FFF]+";
            MatchCollection Matches = Regex.Matches
            (oriText, x, RegexOptions.IgnoreCase);
            StringBuilder sb = new StringBuilder();
            foreach (Match NextMatch in Matches)
            {
                sb.Append(NextMatch.Value);
            }
            return sb.ToString();
        }

        /// <summary>
        /// 正则提取数字
        /// </summary>
        /// <param name="oriText"></param>
        /// <returns></returns>
        public string GetNumbers(string oriText)
        {
            string x = @"\d+";
            MatchCollection Matches = Regex.Matches
            (oriText, x, RegexOptions.IgnoreCase);
            StringBuilder sb = new StringBuilder();
            foreach (Match NextMatch in Matches)
            {
                sb.Append(NextMatch.Value);
            }
            return sb.ToString();
        }



        //获得word文件的文本内容
        public string Doc2Text(string docFileName)
        {
            //实例化COM
            Microsoft.Office.Interop.Word.ApplicationClass wordApp =
                new Microsoft.Office.Interop.Word.ApplicationClass();
            object fileobj = docFileName;
            object nullobj = System.Reflection.Missing.Value;
            //打开指定文件（不同版本的COM参数个数有差异，一般而言除第一个外都用nullobj就行了）
            Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Open(ref fileobj, ref nullobj, ref nullobj,
                ref nullobj, ref nullobj, ref nullobj,

                ref nullobj, ref nullobj, ref nullobj,
                ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj
                );
            //取得doc文件中的文本
            string outText = doc.Content.Text;
            //关闭文件
            doc.Close(ref nullobj, ref nullobj, ref nullobj);
            //关闭COM
            wordApp.Quit(ref nullobj, ref nullobj, ref nullobj);
            //返回
            return outText;
        }

        /// <summary>
        /// 列表转出成DataTable
        /// </summary>
        /// <param name="lst"></param>
        /// <returns></returns>
        public System.Data.DataTable ListToDataTable(List<ApplyInfo> lst)
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            //添加列方法１
            //添加一列列名为id，类型为string
            dt.Columns.Add("姓名", System.Type.GetType("System.String"));//直接为表创建一列//添加列方法//添加一列列名为foldername，类型为string
            dt.Columns.Add("身份证", System.Type.GetType("System.String"));
            dt.Columns.Add("申请日期", System.Type.GetType("System.String"));
            dt.Columns.Add("类型", System.Type.GetType("System.String"));
            dt.Columns.Add("商品", System.Type.GetType("System.String"));
            dt.Columns.Add("个人电话", System.Type.GetType("System.String"));
            dt.Columns.Add("单位名称", System.Type.GetType("System.String"));
            dt.Columns.Add("单位电话", System.Type.GetType("System.String"));
            dt.Columns.Add("与本人关系", System.Type.GetType("System.String"));
            dt.Columns.Add("亲属姓名", System.Type.GetType("System.String"));
            dt.Columns.Add("亲属电话", System.Type.GetType("System.String"));

            //添加行方法
            //添加有数据的行
            foreach (ApplyInfo fo in lst)
            {
                DataRow dr = dt.NewRow();//创建新列
                dr["姓名"] = fo.xingming;//设置列值，通过列名
                dr["身份证"] = fo.shenfenzheng;//设置列值，通过列名
                dr["申请日期"] = fo.shenqingriqi;//设置列值，通过列名
                dr["类型"] = fo.leixing;//设置列值，通过列名
                dr["商品"] = fo.shangping;//设置列值，通过列名
                dr["个人电话"] = fo.shoujihao;//设置列值，通过列名
                dr["单位名称"] = fo.danwei;//设置列值，通过列名
                dr["单位电话"] = fo.danweidianhua;//设置列值，通过列名
                dr["与本人关系"] = fo.guanxi;//设置列值，通过列名
                dr["亲属姓名"] = fo.qinshuxingming;//设置列值，通过列名
                dr["亲属电话"] = fo.qinshushouji;//设置列值，通过列名
                dt.Rows.Add(dr);//想表中添加数据
            }
            return dt;
        }
    }
}
