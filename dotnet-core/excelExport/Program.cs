using System;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using static System.Console;
using System.IO;
using YamlDotNet.Serialization;
using System.Collections;
using System.Collections.Generic;
using McMaster.Extensions.CommandLineUtils;
using System.IO.Enumeration;

namespace excelExport
{
    class Program
    {
		static string[] columnSort = new string[] {
				"custCode",
				"custName",
				"businessId",
				"contractId",
				"certCode",
				"outAssignDate",
				"outAssignCode",
				"outAssignNum",
				"overdueTotalAmount",
				"installmentLoanAmount",
				"monthlyCapital",
				"monthlyInterest",
				"monthFee",
				"currentPenalty",
				"loanAmount",
				"paidAmount",
				"loanPeriod",
				"loanOutPeriod",
				"product",
				"repaymentAccountCode",
				"deliveryDate",
				"outAssignExpireDate",
				"overPeriod",
				"domicileTel",
				"unitTel",
				"mobile",
				"registerAddress",
				"residenceAddress",
				"unitAddress",
				"companyName",
				"immediateFamilyName",
				"immediateFamilyTel",
				"spouseName",
				"spouseTel",
				"contactsName",
				"contactsTel",
				"phoneNum1",
				"callTimes1",
				"phoneNum2",
				"callTimes2",
				"phoneNum3",
				"callTimes3",
				"phoneNum4",
				"callTimes4",
				"phoneNum5",
				"callTimes5",
				"phoneNum6",
				"callTimes6",
				"phoneNum7",
				"callTimes7",
				"phoneNum8",
				"callTimes8",
				"phoneNum9",
				"callTimes9",
				"phoneNum10",
				"callTimes10",
				"phoneNum11",
				"callTimes11",
				"phoneNum12",
				"callTimes12",
				"isApplyExecute",
				"salesCityName",
				"salesSubBranches",
				"subArea",
				"cityName",
				"divisionName",
				"currAssignCompany",
				"litigationStatus",
				"content",
				"collPeriod",
				"collContent"
		};

		static Dictionary<string, string> columnType = new Dictionary<string, string>()
		{
				{"custCode", "string"},
				{"custName", "string"},
				{"businessId", "string"},
				{"contractId", "string"},
				{"certCode", "string"},
				{"outAssignDate", "date"},
				{"outAssignCode", "string"},
				{"outAssignNum", "string"},
				{"overdueTotalAmount", "float"},
				{"installmentLoanAmount", "float"},
				{"monthlyCapital", "float"},
				{"monthlyInterest", "float"},
				{"monthFee", "float"},
				{"currentPenalty", "int"},
				{"loanAmount", "float"},
				{"paidAmount", "float"},
				{"loanPeriod", "int"},
				{"loanOutPeriod", "int"},
				{"product", "string"},
				{"repaymentAccountCode", "string"},
				{"deliveryDate", "date"},
				{"outAssignExpireDate", "date"},
				{"overPeriod", "int"},
				{"domicileTel", "string"},
				{"unitTel", "string"},
				{"mobile", "string"},
				{"registerAddress", "string"},
				{"residenceAddress", "string"},
				{"unitAddress", "string"},
				{"companyName", "string"},
				{"immediateFamilyName", "string"},
				{"immediateFamilyTel", "string"},
				{"spouseName", "string"},
				{"spouseTel", "string"},
				{"contactsName", "string"},
				{"contactsTel", "string"},
				{"phoneNum1", "string"},
				{"callTimes1", "int"},
				{"phoneNum2", "string"},
				{"callTimes2", "int"},
				{"phoneNum3", "string"},
				{"callTimes3", "int"},
				{"phoneNum4", "string"},
				{"callTimes4", "int"},
				{"phoneNum5", "string"},
				{"callTimes5", "int"},
				{"phoneNum6", "string"},
				{"callTimes6", "int"},
				{"phoneNum7", "string"},
				{"callTimes7", "int"},
				{"phoneNum8", "string"},
				{"callTimes8", "int"},
				{"phoneNum9", "string"},
				{"callTimes9", "int"},
				{"phoneNum10", "string"},
				{"callTimes10", "int"},
				{"phoneNum11", "string"},
				{"callTimes11", "int"},
				{"phoneNum12", "string"},
				{"callTimes12", "int"},
				{"isApplyExecute", "string"},
				{"salesCityName", "string"},
				{"salesSubBranches", "string"},
				{"subArea", "string"},
				{"cityName", "string"},
				{"divisionName", "string"},
				{"currAssignCompany", "string"},
				{"litigationStatus", "string"},
				{"content", "string"},
				{"collPeriod", "int"},
				{"collContent", "string"}
		};

		static List<Dictionary<string, string>> ReadData(Config c)
		{
			List<Dictionary<string, string>> data = new List<Dictionary<string, string>>();

			using (MySqlConnection conn = new MySqlConnection(c.database))
			{
				using (MySqlCommand cmd = conn.CreateCommand())
				{
					try
					{
						conn.Open();
						cmd.CommandText = "SELECT * FROM test_data";
						MySqlDataReader myReader = cmd.ExecuteReader();
						while (myReader.Read())
						{
							Dictionary<string, string> dic = new Dictionary<string, string>();
							for (var i = 0; i < columnSort.Length; i++)
							{
								dic.Add(columnSort[i], myReader.GetString(columnSort[i]));
							}
							data.Add(dic);
						}
						myReader.Close();
						myReader.Dispose();
						cmd.Dispose();
					}
					catch (Exception ex)
					{
						WriteLine(ex.Message);
					}
					finally
					{
						conn.Close();
					}
				}
			}
			return data;
		}

		static void ExportExcel(List<Dictionary<string, string>> data)
		{

			ExcelPackage.LicenseContext = LicenseContext.Commercial;
			var file = new FileInfo(Path.Join("./dotnet.xlsx"));
			if (file.Exists)
			{
				file.Delete();
			}

			using (ExcelPackage package = new ExcelPackage(file))
			{
				ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.View.FreezePanes(2, 1);
				//添加头 
				for (var i = 0; i < columnSort.Length; i++)
				{
					worksheet.Cells[1, i + 1].Value = columnSort[i];
				}
				for (var j = 0; j < data.Count; j++)
				{
					Dictionary<string, string> d = data[j];
					for (var k = 0; k < columnSort.Length; k++)
					{
						switch (columnType[columnSort[k]])
						{
							case "int":
								int tmpInt = 0;
								try
								{
									tmpInt = int.Parse(d[columnSort[k]]);
								}
								catch (Exception)
								{
									tmpInt = 0;
								}
								worksheet.Cells[j + 2, k + 1].Value = tmpInt;
								worksheet.Cells[j + 2, k + 1].Style.Numberformat.Format = "0";
								break;
							case "float":
								float tmpFloat = 0.00F;
								try
								{
									tmpFloat = float.Parse(d[columnSort[k]]);
								}
								catch (Exception)
								{
									tmpFloat = 0.00F;
								}
								worksheet.Cells[j + 2, k + 1].Value = tmpFloat;
								worksheet.Cells[j + 2, k + 1].Style.Numberformat.Format = "0.00";
								break;
							default:
								worksheet.Cells[j + 2, k + 1].Value = d[columnSort[k]];
								worksheet.Cells[j + 2, k + 1].Style.Numberformat.Format = "general";
								break;
						}
					}
				}

				package.Save();
			}
		}

		static int Main(string[] args)
        {
			var app = new CommandLineApplication();
			app.HelpOption();

			var optionSubject = app.Option("-c|--conf <conf.yaml>", "conf file format conf.yaml", CommandOptionType.SingleValue);

			app.OnExecute(() =>
			{

				var conf = optionSubject.HasValue()
					? optionSubject.Value()
					: "conf.yaml";

				if (conf.Equals("conf.yaml"))
				{
					conf = Path.Join(Directory.GetCurrentDirectory(), conf);

				}
				var confinfo = new FileInfo(conf);
				if (!confinfo.Exists)
				{
					WriteLine($"{conf} not found!");
					return -1;
				}

				Config c = new Config();
				using (TextReader reader = File.OpenText(conf))
				{
					Deserializer deserializer = new Deserializer();
					c = deserializer.Deserialize<Config>(reader);
				}


				List<Dictionary<string, string>> data = ReadData(c);
				ExportExcel(data);
				return 0;

			});


			try
			{
				var e = app.Execute(args);
				return e;
			}
			catch (Exception e)
			{
				WriteLine(e.Message);
				return -1;
			}
		}
    }
}
