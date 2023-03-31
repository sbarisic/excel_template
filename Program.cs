using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using unvell.ReoGrid;
using System.IO;
using Newtonsoft.Json;

namespace ExcelTemplate {
	class JSONTagSet {
		public string TagName {
			get; set;
		}

		public object TagValue {
			get; set;
		}

		public bool IsFormula {
			get; set;
		}

		public JSONTagSet() {
		}

		public JSONTagSet(string TagName, object TagValue) {
			this.TagName = TagName;
			this.TagValue = TagValue;
			IsFormula = false;
		}

		public override string ToString() {
			return string.Format("{0} = {1}", TagName, TagValue);
		}
	}

	class JSONTemplate {
		public JSONTagSet[] TagSet {
			get; set;
		}

		public JSONTemplate() {
		}

		public JSONTemplate(params JSONTagSet[] TagSet) {
			this.TagSet = TagSet;
		}

		public bool TryFind(string TagName, out JSONTagSet Tag) {
			for (int i = 0; i < TagSet.Length; i++) {
				if (TagSet[i].TagName == TagName) {
					Tag = TagSet[i];
					return true;
				}
			}

			Tag = null;
			return false;
		}
	}

	internal class Program {
		static void Main(string[] args) {
			Console.WriteLine("Starting");


			if (!File.Exists("data/data.json")) {
				Console.WriteLine("data/data.json not found, generated empty example");

				JSONTemplate[] TemplateSample = new JSONTemplate[] {
					new JSONTemplate(new JSONTagSet("TEST", "This is a test value"), new JSONTagSet("SUM", 42)),
					new JSONTemplate(new JSONTagSet("TEST", "This is a test value 2"), new JSONTagSet("SUM", 43))
				};


				File.WriteAllText("data/data.json", JsonConvert.SerializeObject(TemplateSample, Formatting.Indented));

				Console.ReadLine();
				return;
			}

			string DataJSON = File.ReadAllText("data/data.json");
			JSONTemplate[] TemplateValues = JsonConvert.DeserializeObject<JSONTemplate[]>(DataJSON);

			string[] InFiles = Directory.GetFiles("data/in");

			if (!Directory.Exists("data/out"))
				Directory.CreateDirectory("data/out");

			if (TemplateValues.Length < InFiles.Length) {
				Console.WriteLine("Found template values: {0}", TemplateValues.Length);
				Console.WriteLine("Found input files: {0}", InFiles.Length);
				Console.WriteLine("Count does not match, aborting");
				Console.ReadLine();
				return;
			}

			for (int i = 0; i < InFiles.Length; i++) {
				JSONTemplate Template = TemplateValues[i];

				string InFile = InFiles[i];
				string FileName = Path.GetFileNameWithoutExtension(InFile);
				string OutFile = "data/out/" + FileName + "_out.xlsx";

				ProcessInputFile(Template, InFile, OutFile);
			}
		}

		static void ProcessInputFile(JSONTemplate Template, string InputFile, string OutputFile) {
			Console.WriteLine("Processing: {0}", InputFile);

			ReoGridControl Workbook = new ReoGridControl();
			Workbook.Load(InputFile);
			Console.WriteLine("Workbook loaded");

			foreach (var WS in Workbook.Worksheets) {
				ProcessWorksheet(Template, WS);
			}

			Workbook.Save(OutputFile);
		}

		static void ProcessWorksheet(JSONTemplate Template, Worksheet WS) {
			int Rows = WS.Rows;
			int Columns = WS.Columns;

			for (int Y = 0; Y < Rows; Y++) {
				for (int X = 0; X < Columns; X++) {

					object CellData = WS.GetCellData(Y, X);

					if (CellData != null) {
						string CellDataStr = CellData.ToString();

						if (CellDataStr.StartsWith("{") && CellDataStr.EndsWith("}")) {
							string TagName = CellDataStr.Substring(1, CellDataStr.Length - 2);

							if (Template.TryFind(TagName, out JSONTagSet Tag)) {

								if (Tag.IsFormula) {
									WS.SetCellFormula(Y, X, Tag.TagValue.ToString());
								} else {
									WS.SetCellData(Y, X, Tag.TagValue);
								}
							}
						}

					}
				}
			}
		}
	}
}
