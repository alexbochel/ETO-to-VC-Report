using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication2
{
    /// <summary>
    /// This class is a simple structure that carries a string and an integer. 
    /// 
    /// @author: Alexander James Bochel
    /// @version: 8/30/2017
    /// 
    /// </summary>
    public class StringWithCount
    {
        public string str;
        public int count;

        /// <summary>
        /// Set initial count to zero. 
        /// </summary>
        public StringWithCount()
        {
            count = 0;
        }
    }
}
