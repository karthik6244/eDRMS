using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Deloitte.Common.AS2Utilities
{
	public class As2IndexHeader
	{
		public string Version = string.Empty;
		public int Revision = 0;
		public string FolderTitle = string.Empty;
		public int FirstItemIndex = 0;
		public DateTime LastBackupDate;
		public DateTime LastEditDate;
		public DateTime PeriodEndDate;
		public string LongPackName = string.Empty;
		public string PackDir = string.Empty;
		public string PackVersion = string.Empty;
		public short[] PasswordKey = new short[10];
		public int[] Password = new int[10];
		public string DecryptedPassword = string.Empty;
		public string AbkFriendlyName = string.Empty;
	}
}