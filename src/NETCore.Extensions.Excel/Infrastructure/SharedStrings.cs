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
        private Dictionary<string, uint> dic2 = new Dictionary<string, uint>();
        private int nullIndex = 0;

        public bool Exist(string str)
        {
            if (string.IsNullOrEmpty(str))
            {
                return false;
            }
            if (dic2.ContainsKey(str))
                return true;
            return false;
        }

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
                dic2.Add(x.InnerText, index);
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

        public uint _IndexOf(string item)
        {
            return dic2[item];
        }

        public int IndexOf(string item)
        {
            return (int)dic2[item];
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
                dic2.Remove(str);
            }
        }

        public uint _Add(string item)
        {
            lock (this)
            {
                //修复添加null值时，报错的Bug
                if (item == null)
                {
                    nullIndex++;
                    item = "Null";
                }

                if (dic.Count == 0)
                {
                    dic.Add(0, item);
                    dic2.Add(item, 0);
                    return 0;
                }
                else
                {
                    var last = dic.Last().Key;
                    dic.Add(last + 1, item);
                    dic2.Add($"{item}{nullIndex}", last + 1);
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
                    dic2.Add(item, 0);
                }
                else
                {
                    var last = dic.Max(x => x.Key);
                    dic.Add(last + 1, item);
                    dic2.Add(item, last + 1);
                }
            }
        }

        public void Clear()
        {
            dic.Clear();
            dic2.Clear();
        }

        public bool Contains(string item)
        {
            return dic2.ContainsKey(item);
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
    }
}
