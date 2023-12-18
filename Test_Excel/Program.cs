using System;
using OfficeOpenXml;
using System.IO;
using System.Globalization;
using System.Collections.Generic;
using MPCConfig;
using System.Xml.Serialization;
using System.Linq;
using static OfficeOpenXml.ExcelErrorValue;
using System.Reflection;
using System.Security;

namespace Test_Excel
{
    class Program
    {     
        static ExcelPackage pMatrix;
        static ExcelWorksheet wMatrix;
        static ExcelWorksheet wCVs;
        static ExcelWorksheet wMVs;
        static ExcelWorksheet wDVs;
        static ExcelWorksheet wMPC;
        static string fileNameTemplate; 

        static ControllerConfig controllerConfig = new ControllerConfig();

        static ModelTypes modelType = ModelTypes.CONTROL;

        static ModelPriorities modelPriority = ModelPriorities.NORMAL;

        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            Console.WriteLine("Введите имя файла (или полный путь к файлу) шаблона Excel: ");

            fileNameTemplate = Console.ReadLine();

            fileNameTemplate = fileNameTemplate.Trim('"');

            if (fileNameTemplate.Contains(".xlsx"))
            {
                try
                {
                    pMatrix = new ExcelPackage(new FileInfo(fileNameTemplate));

                }
                catch (Exception ex)
                {
                    Console.WriteLine("Ошибка в открытии файла шаблона: " + ex.Message);
                }
            }
            else
            {
                try
                {
                    pMatrix = new ExcelPackage(new FileInfo(fileNameTemplate + ".xlsx"));

                }
                catch (Exception ex)
                {
                    Console.WriteLine("Ошибка в открытии файла шаблона: " + ex.Message);
                }
            }

            wMatrix = pMatrix.Workbook.Worksheets["Matrix"];
            wCVs = pMatrix.Workbook.Worksheets["CV"];
            wMVs = pMatrix.Workbook.Worksheets["MV"];
            wDVs = pMatrix.Workbook.Worksheets["DV"];
            wMPC = pMatrix.Workbook.Worksheets["General"];


            GenerateConfig();

            try
            {
                if (fileNameTemplate.Contains(".xlsx"))
                {
                    fileNameTemplate = fileNameTemplate.Replace(".xlsx", "");

                    XmlSerializer contrexport = new XmlSerializer(typeof(MPCConfig.ControllerConfig));
                    using (FileStream fs = new FileStream(fileNameTemplate + ".xml", FileMode.Create/*FileMode.OpenOrCreate*/))
                    {
                        contrexport.Serialize(fs, controllerConfig);
                    }
                }
                else
                {
                    XmlSerializer contrexport = new XmlSerializer(typeof(MPCConfig.ControllerConfig));
                    using (FileStream fs = new FileStream(fileNameTemplate + ".xml", FileMode.Create/*FileMode.OpenOrCreate*/))
                    {
                        contrexport.Serialize(fs, controllerConfig);
                    }
                }
                
            }catch (Exception ex)
            {
                Console.WriteLine("Ошибка в создании конфигурации: " + ex.Message);
            }
           
            Console.WriteLine("Конфигурация создана успешно, имя файла: " + fileNameTemplate + ".xml");
            Console.ReadLine();
        }



        static void GenerateConfig()
        {
            controllerConfig.CVs = new List<CVConfig>();
            controllerConfig.Models = new List<ModelConfig>();
            controllerConfig.MVs = new List<MVConfig>();
            controllerConfig.DVs = new List<DVConfig>();

            string[] coefs;
            int row_MV = 2;
            int row_CV = 2;
            int row_DV = 2;

            if (wMPC != null)
            {
                if (!string.IsNullOrEmpty((string)wMPC.Cells[2, 2].Value))
                {
                    controllerConfig.ControllerName = Convert.ToString(wMPC.Cells[2, 1].Value);
                }

                if (!string.IsNullOrEmpty((string)wMPC.Cells[2, 2].Value))
                {
                    controllerConfig.ActualStateOPCPath = Convert.ToString(wMPC.Cells[2, 2].Value);
                }

                if (!string.IsNullOrEmpty((string)wMPC.Cells[2, 3].Value))
                {
                    controllerConfig.DesiredStateOPCPath = Convert.ToString(wMPC.Cells[2, 3].Value);
                }

                if (!string.IsNullOrEmpty((string)wMPC.Cells[2, 4].Value))
                {
                    controllerConfig.WatchDogOPCPath = Convert.ToString(wMPC.Cells[2, 4].Value);
                }

                if (!string.IsNullOrEmpty((string)wMPC.Cells[2, 5].Value))
                {
                    controllerConfig.NumberOfCelectedEconFunctionPath = Convert.ToString(wMPC.Cells[2, 5].Value);
                }
            }


            for (int i = 1; i < 50; i++)
            {
                for (int j = 1; j < 50; j++)
                {
                    if (!string.IsNullOrEmpty((string)wMatrix.Cells[i, j].Value))
                    {
                        coefs = Convert.ToString(wMatrix.Cells[i, j].Value).Replace(" ", "").Replace("\n", " ").Replace("=", " ").Replace(".", ",").Split(' ');

                        if (coefs.Length == 1 && coefs[0].Contains("CV") && coefs[0].StartsWith('C'))
                        {
                            if (!string.IsNullOrEmpty((string)wMatrix.Cells[i, j + 1].Value))
                            {
                                controllerConfig.CVs.Add(new CVConfig()
                                {
                                    Name = (string)wMatrix.Cells[i, j + 1].Value,
                                    Description = (string)wMatrix.Cells[i, j + 2].Value ?? string.Empty,
                                    Weigth = 1,
                                    Priority = 1,
                                    EU = string.Empty
                                });
                            }


                            try
                            {
                                int columNum = 2;

                                object parsingValue = wCVs.Cells[row_CV, columNum].Value;

                                SetCVPropertyValue(nameof(CVConfig.POV), controllerConfig.CVs[row_CV - 2], parsingValue);

                                parsingValue = wCVs.Cells[row_CV, ++columNum].Value;

                                SetCVPropertyValue(nameof(CVConfig.LoLimitInput), controllerConfig.CVs[row_CV - 2], parsingValue);

                                parsingValue = wCVs.Cells[row_CV, ++columNum].Value;

                                SetCVPropertyValue(nameof(CVConfig.HiLimitInput), controllerConfig.CVs[row_CV - 2], parsingValue);

                                parsingValue = wCVs.Cells[row_CV, ++columNum].Value;

                                SetCVPropertyValue(nameof(CVConfig.LoLimitEng), controllerConfig.CVs[row_CV - 2], parsingValue);

                                parsingValue = wCVs.Cells[row_CV, ++columNum].Value;

                                SetCVPropertyValue(nameof(CVConfig.HiLimitEng), controllerConfig.CVs[row_CV - 2], parsingValue);

                                parsingValue = wCVs.Cells[row_CV, ++columNum].Value;

                                SetCVPropertyValue(nameof(CVConfig.ActualStateOPCPath), controllerConfig.CVs[row_CV - 2], parsingValue);

                                parsingValue = wCVs.Cells[row_CV, ++columNum].Value;

                                SetCVPropertyValue(nameof(CVConfig.DesiredStateOPCPath), controllerConfig.CVs[row_CV - 2], parsingValue);

                                parsingValue = wCVs.Cells[row_CV, ++columNum].Value;

                                SetCVPropertyValue(nameof(CVConfig.SSValue), controllerConfig.CVs[row_CV - 2], parsingValue);

                                parsingValue = wCVs.Cells[row_CV, ++columNum].Value;

                                SetCVPropertyValue(nameof(CVConfig.TargetRTOCost), controllerConfig.CVs[row_CV - 2], parsingValue);

                                parsingValue = wCVs.Cells[row_CV, ++columNum].Value;

                                SetCVPropertyValue(nameof(CVConfig.TargetRTO), controllerConfig.CVs[row_CV - 2], parsingValue);

                            }
                            catch(Exception ex)
                            {

                                Console.WriteLine(ex.Message);

                            }

                            row_CV++;
                        }
                        else if (coefs.Length == 1 && coefs[0].Contains("MV") && coefs[0].StartsWith('M'))

                        {
                            if (!string.IsNullOrEmpty((string)wMatrix.Cells[i + 1, j].Value))
                            {
                                controllerConfig.MVs.Add(new MVConfig()
                                {
                                    Name = (string)wMatrix.Cells[i + 1, j].Value,
                                    Description = (string)wMatrix.Cells[i + 2, j].Value ?? string.Empty,
                                    Weigth = 1,
                                    EU = string.Empty
                                });
                            }




                            try
                            {
                                int columNum = 2;

                                object parsingValue = wMVs.Cells[row_MV, columNum].Value;

                                SetMVPropertyValue(nameof(MVConfig.PV), controllerConfig.MVs[row_MV - 2], parsingValue);

                                parsingValue = wMVs.Cells[row_MV, ++columNum].Value;

                                SetMVPropertyValue(nameof(MVConfig.SV), controllerConfig.MVs[row_MV - 2], parsingValue);

                                parsingValue = wMVs.Cells[row_MV, ++columNum].Value;

                                SetMVPropertyValue(nameof(MVConfig.RSV), controllerConfig.MVs[row_MV - 2], parsingValue);

                                parsingValue = wMVs.Cells[row_MV, ++columNum].Value;

                                SetMVPropertyValue(nameof(MVConfig.OP), controllerConfig.MVs[row_MV - 2], parsingValue);

                                parsingValue = wMVs.Cells[row_MV, ++columNum].Value;

                                SetMVPropertyValue(nameof(MVConfig.dMVup), controllerConfig.MVs[row_MV - 2], parsingValue);

                                parsingValue = wMVs.Cells[row_MV, ++columNum].Value;

                                SetMVPropertyValue(nameof(MVConfig.dMVdown), controllerConfig.MVs[row_MV - 2], parsingValue);

                                parsingValue = wMVs.Cells[row_MV, ++columNum].Value;

                                SetMVPropertyValue(nameof(MVConfig.dMV), controllerConfig.MVs[row_MV - 2], parsingValue);

                                parsingValue = wMVs.Cells[row_MV, ++columNum].Value;

                                SetMVPropertyValue(nameof(MVConfig.LoLimitInput), controllerConfig.MVs[row_MV - 2], parsingValue);

                                parsingValue = wMVs.Cells[row_MV, ++columNum].Value;

                                SetMVPropertyValue(nameof(MVConfig.HiLimitInput), controllerConfig.MVs[row_MV - 2], parsingValue);

                                parsingValue = wMVs.Cells[row_MV, ++columNum].Value;

                                SetMVPropertyValue(nameof(MVConfig.LoLimitEng), controllerConfig.MVs[row_MV - 2], parsingValue);

                                parsingValue = wMVs.Cells[row_MV, ++columNum].Value;

                                SetMVPropertyValue(nameof(MVConfig.HiLimitEng), controllerConfig.MVs[row_MV - 2], parsingValue);

                                parsingValue = wMVs.Cells[row_MV, ++columNum].Value;

                                SetMVPropertyValue(nameof(MVConfig.ActualStateOPCPath), controllerConfig.MVs[row_MV - 2], parsingValue);

                                parsingValue = wMVs.Cells[row_MV, ++columNum].Value;

                                SetMVPropertyValue(nameof(MVConfig.DesiredStateOPCPath), controllerConfig.MVs[row_MV - 2], parsingValue);

                                parsingValue = wMVs.Cells[row_MV, ++columNum].Value;

                                SetMVPropertyValue(nameof(MVConfig.SSValue), controllerConfig.MVs[row_MV - 2], parsingValue);

                                parsingValue = wMVs.Cells[row_MV, ++columNum].Value;

                                SetMVPropertyValue(nameof(MVConfig.CalculatedValue), controllerConfig.MVs[row_MV - 2], parsingValue);

                                parsingValue = wMVs.Cells[row_MV, ++columNum].Value;

                                SetMVPropertyValue(nameof(MVConfig.OPHi), controllerConfig.MVs[row_MV - 2], parsingValue);

                                parsingValue = wMVs.Cells[row_MV, ++columNum].Value;

                                SetMVPropertyValue(nameof(MVConfig.OPLo), controllerConfig.MVs[row_MV - 2], parsingValue);

                            }
                            catch (Exception ex)
                            {

                                Console.WriteLine(ex.Message);

                            }
                        
                            row_MV++;

                        }
                        else if (coefs.Length == 1 && coefs[0].Contains("DV") && coefs[0].StartsWith('D'))

                        {
                            if (!string.IsNullOrEmpty((string)wMatrix.Cells[i + 1, j].Value))
                            {
                                controllerConfig.DVs.Add(new DVConfig()
                                {
                                    Name = (string)wMatrix.Cells[i + 1, j].Value,
                                    Description = (string)wMatrix.Cells[i + 2, j].Value ?? string.Empty,
                                    EU = string.Empty

                                });
                            }


                            try
                            {

                                int columNum = 2;

                                object parsingValue = wDVs.Cells[row_DV, columNum].Value;

                                SetDVPropertyValue(nameof(DVConfig.Value), controllerConfig.DVs[row_DV - 2], parsingValue);

                                parsingValue = wDVs.Cells[row_DV, ++columNum].Value;

                                SetDVPropertyValue(nameof(DVConfig.ActualStateOPCPath), controllerConfig.DVs[row_DV - 2], parsingValue);

                                parsingValue = wDVs.Cells[row_DV, ++columNum].Value;

                                SetDVPropertyValue(nameof(DVConfig.DesiredStateOPCPath), controllerConfig.DVs[row_DV - 2], parsingValue);

                            }
                            catch (Exception ex)
                            {

                                Console.WriteLine(ex.Message);

                            }

                            row_DV++;
                        }



                        //string a = coefs[1];
                        //double value;
                        //double.TryParse(string.Join("", a.Where(c => char.IsDigit(c))), out value);

                        if (coefs.Contains("G") && coefs.Contains("D") && coefs.Contains("T") && !coefs.Contains("T2"))
                        {
                            try
                            {

                                SetModelProperties(coefs);

                                var typePF = TypesPF.FirstOrder;

                                if (coefs.Contains(nameof(TypesPF.RealInteger)))
                                {
                                    typePF = TypesPF.RealInteger;
                                }                                                           

                                controllerConfig.Models.Add(new ModelConfig()
                                {
                                    Gain = ConvertDoubleCustom(coefs[1]),

                                    T = ConvertDoubleCustom(coefs[5]),

                                    tau = ConvertDoubleCustom(coefs[3]),

                                    cvindex = i - 4,

                                    mvindex = j - 4,

                                    TypePF = typePF,

                                    ModelType = modelType,

                                    Priority = modelPriority
                                });

                            }
                            catch (Exception ex)
                            {

                                Console.WriteLine(ex.Message);

                            }

                        }
                        else if (coefs.Contains("G") && coefs.Contains("D") && !coefs.Contains("T") && !coefs.Contains("T2"))
                        {
                            try
                            {
                                SetModelProperties(coefs);

                                controllerConfig.Models.Add(new ModelConfig()
                                {
                                    Gain = ConvertDoubleCustom(coefs[1]),

                                    tau = ConvertDoubleCustom(coefs[3]),

                                    isIntegrity = true,

                                    cvindex = i - 4,

                                    mvindex = j - 4,

                                    TypePF = TypesPF.Integer,

                                    ModelType = modelType,

                                    Priority = modelPriority
                                });
                            }
                            catch (Exception ex)
                            {

                                Console.WriteLine(ex.Message);

                            }

                        }
                        else if (coefs.Contains("G") && coefs.Contains("D") && coefs.Contains("T") && coefs.Contains("T2"))
                        {
                            try
                            {
                                SetModelProperties(coefs);

                                controllerConfig.Models.Add(new ModelConfig()
                                {
                                    Gain = ConvertDoubleCustom(coefs[1]),

                                    T = ConvertDoubleCustom(coefs[5]),

                                    T2 = ConvertDoubleCustom(coefs[7]),

                                    tau = ConvertDoubleCustom(coefs[3]),

                                    cvindex = i - 4,

                                    mvindex = j - 4,

                                    TypePF = TypesPF.SecondOrder,

                                    ModelType = modelType,

                                    Priority = modelPriority
                                });
                            }
                            catch (Exception ex)
                            {

                                Console.WriteLine(ex.Message);

                            }

                        }
                    }
                }
            }
        }

        static void SetCVPropertyValue(string propertyName, CVConfig cVConfig, object excelValue)
        {
            Type type = typeof(CVConfig);

            var property = type.GetProperty(propertyName);

            if (excelValue is null)
            {
                return;
            }

            string parsExcelValue;

            string oldExcelValue;

            if (!string.IsNullOrEmpty(excelValue.ToString().Replace(',', '.')))
            {
                oldExcelValue = excelValue.ToString();

                parsExcelValue = excelValue.ToString().Replace(',', '.');
            }
            else
            {
                return;
            }


            double localvalue;

            if (property.PropertyType == typeof(ComplexDouble))
            {
                if (double.TryParse(parsExcelValue, NumberStyles.Any, CultureInfo.InvariantCulture, out localvalue))
                {
                    var newValuePropertyValue = new ComplexDouble()
                    {
                        Value = localvalue
                    };

                    property?.SetValue(cVConfig, newValuePropertyValue);
                }
                else
                {
                    parsExcelValue = oldExcelValue;

                    var newValuePropertyValue = new ComplexDouble();

                    if (parsExcelValue.Contains("M:"))
                    {
                        parsExcelValue = parsExcelValue.Replace("M:", "");

                        newValuePropertyValue.ModuleTag = parsExcelValue;
                    }
                    else
                    {
                        newValuePropertyValue.OPCTag = parsExcelValue;
                    }

                    property?.SetValue(cVConfig, newValuePropertyValue);
                }
            }
            else
            {
                if (double.TryParse(parsExcelValue, NumberStyles.Any, CultureInfo.InvariantCulture, out localvalue))
                {
                    property?.SetValue(cVConfig, localvalue);
                }
                else
                {
                    property?.SetValue(cVConfig, parsExcelValue);
                }
            }

        }

        static void SetMVPropertyValue(string propertyName, MVConfig mVConfig, object excelValue)
        {
            Type type = typeof(MVConfig);
            var property = type.GetProperty(propertyName);

            if (excelValue is null)
            {
                return;
            }


            string parsExcelValue;
            string oldExcelValue;
            if (!string.IsNullOrEmpty(excelValue.ToString().Replace(',', '.')))
            {
                oldExcelValue = excelValue.ToString();
                parsExcelValue = excelValue.ToString().Replace(',', '.');
            }
            else
            {
                return;
            }


            double localvalue;
            if (property.PropertyType == typeof(ComplexDouble))
            {
                if (double.TryParse(parsExcelValue, NumberStyles.Any, CultureInfo.InvariantCulture, out localvalue))
                {
                    var newValuePropertyValue = new ComplexDouble()
                    {
                        Value = localvalue
                    };

                    property?.SetValue(mVConfig, newValuePropertyValue);
                }
                else
                {
                    parsExcelValue = oldExcelValue;

                    var newValuePropertyValue = new ComplexDouble();

                    if (parsExcelValue.Contains("M:"))
                    {
                        newValuePropertyValue.ModuleTag = parsExcelValue;
                    }
                    else
                    {
                        newValuePropertyValue.OPCTag = parsExcelValue;
                    }

                    property?.SetValue(mVConfig, newValuePropertyValue);

                }
            }
            else
            {
                if (double.TryParse(parsExcelValue, NumberStyles.Any, CultureInfo.InvariantCulture, out localvalue))
                {
                    property?.SetValue(mVConfig, localvalue);
                }
                else
                {
                    property?.SetValue(mVConfig, parsExcelValue);
                }
            }

        }

        static void SetDVPropertyValue(string propertyName, DVConfig dVConfig, object excelValue)
        {
            Type type = typeof(DVConfig);

            var property = type.GetProperty(propertyName);

            if (excelValue is null)
            {
                return;
            }


            string parsExcelValue;

            string oldExcelValue;

            if (!string.IsNullOrEmpty(excelValue.ToString().Replace(',', '.')))
            {
                oldExcelValue = excelValue.ToString();
                parsExcelValue = excelValue.ToString().Replace(',', '.');
            }
            else
            {
                return;
            }

            double localvalue;

            if (property.PropertyType == typeof(ComplexDouble))
            {
                if (double.TryParse(parsExcelValue, NumberStyles.Any, CultureInfo.InvariantCulture, out localvalue))
                {
                    var newValuePropertyValue = new ComplexDouble()
                    {
                        Value = localvalue
                    };

                    property?.SetValue(dVConfig, newValuePropertyValue);
                }
                else
                {
                    parsExcelValue = oldExcelValue;

                    var newValuePropertyValue = new ComplexDouble();

                    if (parsExcelValue.Contains("M:"))
                    {
                        newValuePropertyValue.ModuleTag = parsExcelValue;
                    }
                    else
                    {
                        newValuePropertyValue.OPCTag = parsExcelValue;
                    }

                    

                    property?.SetValue(dVConfig, newValuePropertyValue);

                }
            }
            else
            {
                if (double.TryParse(parsExcelValue, NumberStyles.Any, CultureInfo.InvariantCulture, out localvalue))
                {
                    property?.SetValue(dVConfig, localvalue);
                }
                else
                {
                    property?.SetValue(dVConfig, parsExcelValue);
                }
            }

        }

        static double ConvertDoubleCustom(string value)
        {
            string parsExcelValue;

            if (!string.IsNullOrEmpty(value.ToString().Replace(',', '.')))
            {
                parsExcelValue = value.ToString().Replace(',', '.');
            }
            else
            {
                return double.NaN;
            }

            if (double.TryParse(parsExcelValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double localvalue))
            {
                return localvalue;
            }
            else
            {
                return double.NaN;
            }
        }

        static void SetModelProperties(string[] coefs)
        {
            if (coefs.Contains(nameof(ModelTypes.PREDICTION)))
            {
                modelType = ModelTypes.PREDICTION;

            }else if (coefs.Contains(nameof(ModelTypes.CONTROL)))
            {
                modelType = ModelTypes.CONTROL;
            }

            if (coefs.Contains(nameof(ModelPriorities.LOW)))
            {
                modelPriority = ModelPriorities.LOW;

            }
            else if (coefs.Contains(nameof(ModelPriorities.LOWEST)))
            {
                modelPriority = ModelPriorities.LOWEST;

            }else if (coefs.Contains(nameof(ModelPriorities.NORMAL)))
            {
                modelPriority = ModelPriorities.NORMAL;
            }
        }
    }
}
