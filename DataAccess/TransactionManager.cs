using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Transactions;

namespace Common.SqlHelper
{
    public class TransactionManager 
    {
        public void Wrapper(Action method)
        {
            using (var ts = new TransactionScope())
            {
                var retires = 3;
                var succeeded = false;
                while (!succeeded)
                {
                    try
                    {
                        method();
                        ts.Complete();
                        succeeded = true;
                    }
                    catch (Exception ex)
                    {
                        if (retires >= 0)
                            retires--;
                        else
                        {                            
                            throw ex;
                        }
                    }
                }
            }
        }
    }
}
