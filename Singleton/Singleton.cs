using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SampleCode.Toys
{
    public class Singleton
    {
        // 因为构造函数是私有的，所以需要使用lambda
        private static readonly Lazy<Singleton> _instance = new Lazy<Singleton>(() => new Singleton());
        // 由于Lazy<T>默认的设置就是线程安全，所以不设置也是有效的
        // new Lazy<Singleton>(() => new Singleton(), LazyThreadSafetyMode.ExecutionAndPublication);

        private Singleton()
        {
        }

        public static Singleton Instance
        {
            get
            {
                return _instance.Value;
            }
        }
    }
}
