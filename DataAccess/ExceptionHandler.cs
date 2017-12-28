using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Common.SqlHelper
{
    public class ExceptionHandler
    {
        public void Wrapper(Action method)
        {
            try
            {
                method();
            }
            catch (Exception ex)
            {
                //TODO  log or something for write
                throw ex;
            }
        }
    }
}
