/*************************************************************************************
     * CLR版本：       4.0.30319.17929
     * 类 名 称：       ApplyInfo
     * 机器名称：       2016-20170206KH
     * 命名空间：       SwitchPdf
     * 文 件 名：       ApplyInfo
     * 创建时间：       2017/6/2 13:53:08
     * 作    者：       LIU FANG
     * 说   明：。。。。。
     * 修改时间：
     * 修 改 人：
    *************************************************************************************/


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SwitchPdf
{
    /// <summary>
    ///  客户申请信息
    /// </summary>
    public class ApplyInfo
    {
        /// <summary>
        ///姓名
        /// </summary>
        public string xingming { get; set; }
        /// <summary>
        /// 身份证
        /// </summary>
        public string shenfenzheng { get; set; }
        /// <summary>
        /// 申请日期
        /// </summary>
        public string shenqingriqi { get; set; }
        /// <summary>
        /// 申请类型
        /// </summary>
        public string leixing { get; set; }
        /// <summary>
        /// 商品
        /// </summary>
        public string shangping { set; get; }
        /// <summary>
        /// 单位
        /// </summary>
        public string danwei { get; set; }
        /// <summary>
        /// 单位电话
        /// </summary>
        public string danweidianhua { get; set; }
        /// <summary>
        /// 个人手机号
        /// </summary>
        public string shoujihao { get; set; }
        /// <summary>
        /// 与本人关系
        /// </summary>
        public string guanxi { get; set; }
        /// <summary>
        /// 亲属姓名
        /// </summary>

        public string qinshuxingming { get; set; }

        /// <summary>
        /// 亲属手机
        /// </summary>
        public string qinshushouji { get; set; }



    }
}
