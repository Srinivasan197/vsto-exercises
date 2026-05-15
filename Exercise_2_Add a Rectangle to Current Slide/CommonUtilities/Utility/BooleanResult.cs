using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExampleVSTO.CommonUtilities.Utility
{
    public class BooleanResult<T>
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public T Result { get; set; }
        public static BooleanResult<T> SuccessResult(T result)
        {
            return new BooleanResult<T>
            {
                Success = true,
                Result = result
            };
        }
        public static BooleanResult<T> FailResult(string error)
        {
            return new BooleanResult<T>
            {
                Success = false,
                Message = error
            };
        }
    }
}
