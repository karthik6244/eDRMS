using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ICSharpCode.SharpZipLib.Zip;

namespace Deloitte.Common.AS2Utilities
{
	/// <summary>
	/// Class for manipulating AS/2 Index files
	/// 
	/// Brian P. Mueller
	/// Deloitte Services LP
	/// July 31st, 2009
	/// </summary>
	public class As2IndexTools
	{
		private const string As2IndexFileName = "index.sav";
		private const int As2IndexFileHeaderLength = 648;
		private const int As2IndexFileSegmentLength = 632;

		/// <summary>
		/// Public method to initiate extraction and handle the ZIP file.
		/// </summary>
		/// <param name="AS2FilePath">Path to the AS/2 file that we want to extract the index information from.</param>
		/// <returns>Populated AS2Index object with all data from the index file</returns>
		public As2Index ExtractAs2IndexInfo(string AS2FilePath)
		{
			ZipEntry indexFileEntry;
			As2Index indexContents = null;

			ZipFile zipFile = new ZipFile(AS2FilePath);
			indexFileEntry = zipFile.GetEntry(As2IndexFileName);

			string zipComment = zipFile.ZipFileComment;

			// Extract the Index.sav file to memory
			MemoryStream indexStream = null;
			MemoryStream headerStream = null;

			using (Stream zipStream = zipFile.GetInputStream(indexFileEntry))
			{
				byte[] streamSizingBuffer = new byte[indexFileEntry.Size];

				// Make two copies of the stream since Header and Index will be processed separately.
				// We need copies so that we can make the stream seekable and release the underlying file.
				indexStream = new MemoryStream(streamSizingBuffer,true);
				headerStream = new MemoryStream(streamSizingBuffer, true);

				StreamReader zipReader = new StreamReader(zipStream);
				int readCount;
				byte[] copyBuffer = new byte[1024];
				while ((readCount = zipStream.Read(copyBuffer, 0, 1024)) != 0)
				{
					indexStream.Write(copyBuffer, 0, readCount);
					headerStream.Write(copyBuffer, 0, readCount);
				}
				indexStream.Position = 0;
				headerStream.Position = 0;
			}

			indexContents = GetAs2Properties(headerStream,indexStream);

			ProcessHierarchy(indexContents);

			indexContents.Header.AbkFriendlyName = zipComment.Substring(zipComment.LastIndexOf(";") + 1);

			return indexContents;
		}

		/// <summary>
		/// Process the parent/child tree of an AS/2 index to set the indent/tree levels
		/// </summary>
		/// <param name="index">AS/2 Index to process</param>
		private void ProcessHierarchy(As2Index index)
		{
			int previousIndex = 0;
			int previousParentIndex = 0;
			int indentLevel = 0;

			foreach (AS2Utilities.As2IndexRecord record in index.Records)
			{
				if (record.Parent == previousIndex && record.Parent > 0)
				{
					indentLevel++;
				}
				else if (record.Parent != previousParentIndex)
				{
					indentLevel--;
				}

				record.TreeLevel = indentLevel;

				previousIndex = record.Index;
				previousParentIndex = record.Parent;
			}
		}

		/// <summary>
		/// Wrapper method to drive the extraction of the pieces of the index file.
		/// </summary>
		/// <param name="headerStream">Stream containing the decompressed Index file data</param>
		/// <param name="indexStream">Stream containing the decompressed Index file data</param>
		/// <returns>Populated AS2Index object with all data from the index file</returns>
		private As2Index GetAs2Properties(MemoryStream headerStream, MemoryStream indexStream)
		{
			As2IndexHeader packageHeader = GetHeader(headerStream);
			LinkedList<As2IndexRecord> fileIndex = GetFileList(indexStream, packageHeader);

			As2Index completedProperties = new As2Index();

			completedProperties.Header = packageHeader;
			completedProperties.Records = fileIndex;

			return completedProperties;
		}

		/// <summary>
		/// Extract the Header data from the AS2 index file.
		/// </summary>
		/// <param name="headerStream">Stream containing the decompressed Index file data</param>
		/// <returns>Decoded Header data</returns>
		private As2IndexHeader GetHeader(MemoryStream headerStream)
		{
			BinaryReader readBinary = null;
			try
			{
				As2IndexHeader header = new As2IndexHeader();

				readBinary = new BinaryReader(headerStream);
				header.Version = Encoding.Unicode.GetString(readBinary.ReadBytes(12)).TrimEnd('\0');
				header.Revision = BitConverter.ToInt32(readBinary.ReadBytes(sizeof(int)),0);
				header.FolderTitle = Encoding.Unicode.GetString(readBinary.ReadBytes(164)).TrimEnd('\0');
				header.FirstItemIndex = BitConverter.ToInt32(readBinary.ReadBytes(4), 0);
				header.LastBackupDate = ConvertVB6DateToDateTime(readBinary.ReadBytes(8));
				header.LastEditDate = ConvertVB6DateToDateTime(readBinary.ReadBytes(8));
				header.PeriodEndDate = ConvertVB6DateToDateTime(readBinary.ReadBytes(8));
				header.LongPackName = Encoding.Unicode.GetString(readBinary.ReadBytes(162)).TrimEnd('\0');
				header.PackDir = Encoding.Unicode.GetString(readBinary.ReadBytes(162)).TrimEnd('\0');
				header.PackVersion = Encoding.Unicode.GetString(readBinary.ReadBytes(12)).TrimEnd('\0');

				for (int i=0; i < 10; i++)
				{
					header.PasswordKey[i] = BitConverter.ToInt16(readBinary.ReadBytes(sizeof(short)), 0);
				}

				for (int i=0; i < 10; i++)
				{
					header.Password[i] = BitConverter.ToInt16(readBinary.ReadBytes(sizeof(int)), 0);
				}

				readBinary.Close();
				readBinary = null;

				//Decrypt password if there is one
				char decryptedCharacter;
				for (int i = 0; i < 10; i++)
				{
					if (header.Password[i] != 0)
					{
						decryptedCharacter = (char)(header.Password[i] / header.PasswordKey[i]);
						header.DecryptedPassword += decryptedCharacter;
					}
				}

				return header;
			}
			finally
			{
				if (readBinary != null)
				{
					readBinary.Close();
				}
			}
		}

		/// <summary>
		/// Extract the list of files and their attributes from the AS2 index file.
		/// </summary>
		/// <param name="zipStream">Stream containing the decompressed Index file data</param>
		/// <param name="packageHeader">Object containing the decoded Index file header data</param>
		/// <returns>Decoded file list information</returns>
		private LinkedList<As2IndexRecord> GetFileList(MemoryStream indexStream, As2IndexHeader packageHeader)
		{
			BinaryReader readBinary = null;

			try
			{
				LinkedList<As2IndexRecord> fileIndexCollection = new LinkedList<As2IndexRecord>();
				
				readBinary = new BinaryReader(indexStream);
				int nextItemIndex = packageHeader.FirstItemIndex;

				while (nextItemIndex > 0)
				{
					int startByte = As2IndexFileHeaderLength + (nextItemIndex - 1) * As2IndexFileSegmentLength;
					indexStream.Seek(startByte, SeekOrigin.Begin);

					As2IndexRecord record = new As2IndexRecord();

					record.Title = Encoding.Unicode.GetString(readBinary.ReadBytes(164)).TrimEnd('\0');
					record.Index = nextItemIndex;
					record.Segment = BitConverter.ToInt32(readBinary.ReadBytes(sizeof(int)), 0);
					record.Parent = BitConverter.ToInt32(readBinary.ReadBytes(sizeof(int)), 0);
					record.NextItemIndex = BitConverter.ToInt32(readBinary.ReadBytes(sizeof(int)), 0);
					record.UID = Encoding.Unicode.GetString(readBinary.ReadBytes(76)).TrimEnd('\0');
					record.ItemType = BitConverter.ToInt32(readBinary.ReadBytes(sizeof(int)), 0);
					record.DocumentType = Encoding.Unicode.GetString(readBinary.ReadBytes(18)).TrimEnd('\0');
					record.Reference = Encoding.Unicode.GetString(readBinary.ReadBytes(22)).TrimEnd('\0');
					record.IsMaster = BitConverter.ToInt32(readBinary.ReadBytes(sizeof(int)), 0);

					for (int i = 0; i < 4; i++)
					{
						record.PreparedInitials[i] = Encoding.Unicode.GetString(readBinary.ReadBytes(22)).TrimEnd('\0');
					}

					for (int i = 0; i < 4; i++)
					{
						record.ReviewInitials[i] = Encoding.Unicode.GetString(readBinary.ReadBytes(22)).TrimEnd('\0');
					}

					for (int i = 0; i < 4; i++)
					{
						record.Offset[i] = readBinary.ReadByte();
					}

					for (int i = 0; i < 4; i++)
					{
						record.PreparedDates[i] = ConvertVB6DateToDateTime(readBinary.ReadBytes(8));
					}

					for (int i = 0; i < 4; i++)
					{
						record.ReviewedDates[i] = ConvertVB6DateToDateTime(readBinary.ReadBytes(8));
					}

					record.IsAttentionManual = BitConverter.ToInt32(readBinary.ReadBytes(sizeof(int)), 0);
					record.IsAttentionAuto = BitConverter.ToInt32(readBinary.ReadBytes(sizeof(int)), 0);
					record.NumberOfOpenNotes = BitConverter.ToInt32(readBinary.ReadBytes(sizeof(int)), 0);
					record.NumberOfClosedNotes = BitConverter.ToInt32(readBinary.ReadBytes(sizeof(int)), 0);
					record.IsRecentlyFiled = BitConverter.ToInt32(readBinary.ReadBytes(sizeof(int)), 0);
					record.DefaultReference = Encoding.Unicode.GetString(readBinary.ReadBytes(22)).TrimEnd('\0');

					fileIndexCollection.AddLast(record);
					nextItemIndex = record.NextItemIndex;
				}
				
				readBinary.Close();
				readBinary = null;

				return fileIndexCollection;
			}
			finally
			{
				if (readBinary != null)
				{
					readBinary.Close();
				}
			}
		}

		/// <summary>
		/// Convert a set of 8 bytes from the Visual Basic 6.0 date format (which is stored in an IEEE 64-bit floating point value)
		/// to a .Net DateTime data type. 
		/// </summary>
		/// <param name="vb6Date">The 8 bytes comprising the VB6 Date type</param>
		/// <returns>The .Net DateTime value converted from the VB6 Date type</returns>
		private DateTime ConvertVB6DateToDateTime(byte[] vb6Date)
		{
			return DateTime.FromOADate(BitConverter.ToDouble(vb6Date, 0));
		}
	}
}