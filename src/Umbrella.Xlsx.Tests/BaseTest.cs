using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Umbrella.Xlsx.Tests
{
	[TestClass]
	public abstract class BaseTest
	{
		protected List<string> _SavedFiles;

		public abstract bool UseOutputFolder { get; }
		public string OutputFolder { get; private set; }

		private TestContext m_testContext;

		public TestContext TestContext
		{
			get { return m_testContext; }
			set { m_testContext = value; }
		}

		[TestInitialize]
		public void TestDataInitiliaze()
		{
			OutputFolder = Path.Combine(System.Environment.CurrentDirectory, @"..\..\OutputFiles");
			this._SavedFiles = new List<string>();

			// calculate output folder
			PrepareTestData();
		}

		[TestCleanup]
		public void TestDataCleanUp()
		{
			CleanUpTestData();
		}

		protected virtual void PrepareTestData()
		{

		}

		protected virtual void CleanUpTestData()
		{
			if (UseOutputFolder && _SavedFiles != null)
				_SavedFiles.ForEach(x => File.Delete(Path.Combine(this.OutputFolder, x)));
		}
	}
}
