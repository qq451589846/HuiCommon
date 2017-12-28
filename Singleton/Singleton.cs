using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SampleCode.Toys
{
    public class Singleton
    {
        // ��Ϊ���캯����˽�еģ�������Ҫʹ��lambda
        private static readonly Lazy<Singleton> _instance = new Lazy<Singleton>(() => new Singleton());
        // ����Lazy<T>Ĭ�ϵ����þ����̰߳�ȫ�����Բ�����Ҳ����Ч��
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
