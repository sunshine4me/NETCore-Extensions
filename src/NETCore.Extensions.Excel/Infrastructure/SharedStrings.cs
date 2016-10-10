using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;

namespace NETCore.Extensions.Excel.Infrastructure
{
    public class SharedStrings : IList<string>, IDisposable
    {
        private string xmlSource;

        private Dictionary<uint, string> dic = new Dictionary<uint, string>();



        public int Count
        {
            get
            {
                return dic.Count;
            }
        }

        public bool IsReadOnly
        {
            get
            {
                return true;
            }
        }

        public string this[int index]
        {
            get
            {
                return this[Convert.ToUInt32(index)];
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        public SharedStrings(string XmlSource)
        {
            xmlSource = XmlSource;
            var xd = new XmlDocument();
            xd.LoadXml(xmlSource.Replace("standalone=\"true\"", "standalone=\"yes\""));
            var t = xd.GetElementsByTagName("t");
            uint i = 0;
            foreach (XmlNode x in t)
            {
                var index = i++;
                dic.Add(index, x.InnerText);
            }
            xmlSource = null;
            GC.Collect();
        }

        public string this[uint index]
        {
            get
            {
                return dic[index];
            }
            set
            {
                dic[index] = value;
            }
        }

        public void Dispose()
        {
            dic.Clear();
            GC.Collect();
        }



        public void Insert(int index, string item)
        {
            throw new NotImplementedException();
        }

        public void RemoveAt(int index)
        {
            lock(this)
            {
                var str = dic[Convert.ToUInt32(index)];
                dic.Remove(Convert.ToUInt32(index));
            }
        }

        public uint _Add(string item)
        {
            lock (this)
            {
                //修复添加null值时，报错的Bug
                if (item == null)
                {
                    item = "Null";
                }

                if (dic.Count == 0)
                {
                    dic.Add(0, item);
                    return 0;
                }
                else
                {
                    var last = dic.Last().Key;
                    dic.Add(last + 1, item);
                    return last + 1;
                }
            }
        }

        public void Add(string item)
        {
            lock(this)
            {
                if (dic.Count == 0)
                {
                    dic.Add(0, item);
                }
                else
                {
                    var last = dic.Max(x => x.Key);
                    dic.Add(last + 1, item);
                }
            }
        }

        public void Clear()
        {
            dic.Clear();
        }

        public void CopyTo(string[] array, int arrayIndex)
        {
            throw new NotImplementedException();
        }

        public bool Remove(string item)
        {
            lock(this)
            {
                var keys = dic.Where(x => x.Value == item)
                    .Select(x => x.Key)
                    .ToList();
                if (keys.Count == 0)
                    return false;
                foreach (var x in keys)
                    dic.Remove(x);
                return true;
            }
        }

        public IEnumerator<string> GetEnumerator()
        {
            return dic.Select(x => x.Value)
                .ToList()
                .GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public int IndexOf(string item) {
            return  -1;
        }

        public bool Contains(string item) {
            var d = dic.FirstOrDefault(t => t.Value == item);
            if (d.Value != item)
                return false;
            else
                return true;
        }
    }
}
