using Ionic.Zip;
using MunicipalServiceFrguCodesHelper.Properties;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace MunicipalServiceFrguCodesHelper
{
    class Program
    {
        private const int ServiceCodeColumn = 1;
        private const int ServiceNameColumn = 2;
        private const int FirstColumn = 3;

        private const int departmentNameRow = 1;
        private const int departmentCodeRow = 2;
        private const int FirstRow = 3;

        private const string isn_classif_id = "fa101c64-0e12-4ee7-ba9e-3c5b7c263d90";
        private const string isn_classif_name = "Виды заявлений ДСР";

        static void Main(string[] args)
        {
            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(OnAssemblyResolve);
            try
            {
                Console.WriteLine("Укажите полный путь к эксель-файлу, содержащий таблицу кодов ФРГУ муниципальных услуг (следует указать весь путь вместе с расширением файла):");
                var path = new FileInfo(Console.ReadLine());
                if (path.Exists)
                {
                    var result = LoadFile(path);

                    if (!string.IsNullOrEmpty(result))
                    {
                        Save(result, path);
                    }
                }
                else
                {
                    Console.WriteLine($"По указанному пути ({path.FullName}) файл не существует!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
            }

            Console.ReadKey();
        }

        private static Assembly OnAssemblyResolve(object sender, ResolveEventArgs args)
        {

            if (args.Name.Contains("DotNetZip"))
                return Assembly.Load(Resources.DotNetZip);

            if (args.Name.Contains("EPPlus"))
                return Assembly.Load(Resources.EPPlus);

            return null;
        }

        private static string LoadFile(FileInfo file)
        {
            using (var package = new ExcelPackage(file))
            {
                Console.WriteLine($"Файл {file.FullName} открыт.");
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet is null)
                {
                    Console.WriteLine($"Файл {file.FullName} не содержит ни одного листа!");
                    return string.Empty;
                }

                return Parse(worksheet);
            }
        }

        private static string Parse(ExcelWorksheet worksheet)
        {
            Console.WriteLine($"Начат разбор листа {worksheet.Name}.");

            var classifInfoCollection = new List<ClassifInfo>();

            if (worksheet.Dimension.End.Column < FirstColumn)
                throw new Exception($"Лист {worksheet.Name} содержит число столбцов {worksheet.Dimension.End.Column}, что меньше минимально допустимого!");
            if (worksheet.Dimension.End.Row < FirstRow)
                throw new Exception($"Лист {worksheet.Name} содержит число строк {worksheet.Dimension.End.Row}, что меньше минимально допустимого!");

            for (int column = FirstColumn; column <= worksheet.Dimension.End.Column; column++)
            {
                var departmentCode = worksheet.Cells[departmentCodeRow, column].Value?.ToString();
                if (string.IsNullOrEmpty(departmentCode))
                {
                    Console.WriteLine($"Не определен 'Код ведомства', адрес ячейки [{departmentCodeRow}:{column}]");
                    continue;
                }

                var departmentName = worksheet.Cells[departmentNameRow, column].Value?.ToString();
                if (string.IsNullOrEmpty(departmentName))
                {
                    Console.WriteLine($"Не определен 'Наименование ведомства', адрес ячейки [{departmentNameRow}:{column}]");
                    continue;
                }

                for (int row = FirstRow; row <= worksheet.Dimension.End.Row; row++)
                {
                    var serviceCode = worksheet.Cells[row, ServiceCodeColumn].Value?.ToString();
                    if (string.IsNullOrEmpty(serviceCode))
                    {
                        Console.WriteLine($"Не определен 'Код услуги', адрес ячейки [{row}:{ServiceCodeColumn}]");
                        continue;
                    }
                    var serviceName = worksheet.Cells[row, ServiceNameColumn].Value?.ToString();
                    if (string.IsNullOrEmpty(serviceName))
                    {
                        Console.WriteLine($"Не определен 'Наименование услуги', адрес ячейки [{row}:{ServiceNameColumn}]");
                        continue;
                    }

                    var codeFrgu = worksheet.Cells[row, column]?.Text.Trim().ToUpper();
                    if (!string.IsNullOrEmpty(codeFrgu) && !(codeFrgu.StartsWith("ВПР(") || codeFrgu.StartsWith("VLOOKUP(")))
                    {
                        classifInfoCollection.Add(new ClassifInfo()
                        {
                            CodeFrgu = codeFrgu,
                            DepartmentCode = departmentCode,
                            DepartmentName = departmentName,
                            ServiceCode = serviceCode,
                            ServiceName = serviceName
                        });
                    }
                }
            }

            if (!classifInfoCollection.Any())
            {
                Console.WriteLine("Данные отсутствуют.");
                return string.Empty;
            }

            Console.WriteLine($"Сформировано данных {classifInfoCollection.Count} шт.");
            return XmlContentBuild(classifInfoCollection);
        }

        private static void Save(string content, FileInfo path)
        {
            var _path = Path.Combine(path.DirectoryName, $"{Path.GetFileNameWithoutExtension(path.Name)}_{DateTime.Now.ToString("yyyy.MM.dd-HH.mm.ss")}.zip");
            var filename = $"Classif_{isn_classif_id}.xml";
            using (var zipFile = new ZipFile())
            {
                zipFile.AddEntry(filename, content, Encoding.UTF8);
                zipFile.Save(_path);
            }
            Console.WriteLine($"Конечный файл сформирован и сохранен: {_path}");
        }

        private static string XmlContentBuild(IList<ClassifInfo> collection)
        {
            var sb = new StringBuilder();
            sb.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            sb.AppendLine("<ClassifCard xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");
            sb.AppendLine("<TableId>custom_classif</TableId>");
            sb.AppendLine("<Custom>");
            sb.AppendLine($"<isn_classif>{isn_classif_id}</isn_classif>");
            sb.AppendLine($"<name>{isn_classif_name}</name>");
            sb.AppendLine("<field_count>5</field_count>");
            sb.AppendLine("<fields>");
            sb.AppendLine("<name0>Код ведомства</name0>");
            sb.AppendLine("<name1>Наименование ведомства</name1>");
            sb.AppendLine("<name2>Код услуги</name2>");
            sb.AppendLine("<name3>Наименование услуги</name3>");
            sb.AppendLine("<name4>Код ФРГУ</name4>");
            sb.AppendLine("</fields>");
            sb.AppendLine("<is_hierarchical>false</is_hierarchical>");
            sb.AppendLine("</Custom>");

            sb.AppendLine("<CustomRows>");
            foreach (var item in collection)
            {
                sb.AppendLine("<custom_classif_row>");
                sb.AppendLine($"<isn_node>{Guid.NewGuid()}</isn_node>");
                sb.AppendLine("<isn_parent_node xsi:nil=\"true\" />");
                sb.AppendLine("<is_parent>false</is_parent>");
                sb.AppendLine($"<field0>{item.DepartmentCode}</field0>");
                sb.AppendLine($"<field1>{item.DepartmentName}</field1>");
                sb.AppendLine($"<field2>{item.ServiceCode}</field2>");
                sb.AppendLine($"<field3>{item.ServiceName}</field3>");
                sb.AppendLine($"<field4>{item.CodeFrgu}</field4>");
                sb.AppendLine("</custom_classif_row>");
            }
            sb.AppendLine("</CustomRows>");
            sb.AppendLine("</ClassifCard>");

            return sb.ToString();
        }

        private class ClassifInfo
        {
            /// <summary> Код ФРГУ </summary>
            public string CodeFrgu { get; set; }

            /// <summary> Код услуги </summary>
            public string ServiceCode { get; set; }

            /// <summary> Наименование услуги </summary>
            public string ServiceName { get; set; }

            /// <summary> Код ведомства </summary>
            public string DepartmentCode { get; set; }

            /// <summary> Наименование ведомства </summary>
            public string DepartmentName { get; set; }
        }
    }
}
