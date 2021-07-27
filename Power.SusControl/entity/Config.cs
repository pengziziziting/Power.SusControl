using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Power.Controls.SystemCESE.entity
{
    /// <summary>
    /// 配置文件实体类
    /// </summary>
    public class Config
    {
        public string keyword { get; set; }
        public string fields { get; set; }

        public string filter { get; set; }
    }


    public class ConfigChildren
    {
        public string KeyWord { get; set; }
        public Hashtable fields { get; set; }
        public Hashtable filter { get; set; }
        public string miniid { get; set; }

        public string KeyWordType { get; set; }
        public string sort { get; set; }

        public string swhere { get; set; }

        public List<ConfigChildren> children;
    }
}
