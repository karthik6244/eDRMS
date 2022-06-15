using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Deloitte.Common.AS2Utilities
{
	public class As2IndexRecord
	{
		public string Title = string.Empty;
		public int Index = 0;
		public int TreeLevel = 0;
		public int Segment = 0;
		public int Parent = 0;
		public int NextItemIndex = 0;
		public string UID = string.Empty;
		public int ItemType = 0;
		public string DocumentType = string.Empty;
		public string Reference = string.Empty;
		public int IsMaster = 0;
		public string[] PreparedInitials = new string[4];
		public string[] ReviewInitials = new string[4];
		public byte[] Offset = new byte[4];
		public DateTime[] PreparedDates = new DateTime[4];
		public DateTime[] ReviewedDates = new DateTime[4];
		public int IsAttentionManual = 0;
		public int IsAttentionAuto = 0;
		public int NumberOfOpenNotes = 0;
		public int NumberOfClosedNotes = 0;
		public int IsRecentlyFiled = 0;
		public string DefaultReference = string.Empty;
	}
}