using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Deloitte.Common.AS2Utilities
{
	public class As2Index
	{
		public As2IndexHeader Header = null;
		public LinkedList<As2IndexRecord> Records = null;
	}
}