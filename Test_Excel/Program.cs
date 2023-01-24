using System;
using OfficeOpenXml;
using System.IO;
using System.Globalization;
using System.Collections.Generic;
using MPCConfig;
using System.Xml.Serialization;
using System.Linq;

namespace Test_Excel
{
    class Program
    {
        static List<string> Lines;
        static List<ReadModel> readModels = new List<ReadModel>();

        static List<Model> models = new List<Model>();

        static List<CV> CVs = new List<CV>();
        static List<MV> MVs = new List<MV>();
        static ExcelPackage pMatrix;
        static ExcelWorksheet wMatrix;
        static ExcelWorksheet wCVs;
        static ExcelWorksheet wMVs;
        static ExcelWorksheet wDVs;
        static ExcelWorksheet wMPC;

        static ExcelWorksheet wApplicaion;

        static ControllerConfig controllerConfig = new ControllerConfig();

        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            pMatrix = new ExcelPackage(new FileInfo("Configuration_Data.xlsx"));

            wMatrix = pMatrix.Workbook.Worksheets["Matrix"];
            wCVs = pMatrix.Workbook.Worksheets["CV"];
            wMVs = pMatrix.Workbook.Worksheets["MV"];
            wDVs = pMatrix.Workbook.Worksheets["DV"];
            wMPC = pMatrix.Workbook.Worksheets["General"];


            GenerateConfig();

            XmlSerializer contrexport = new XmlSerializer(typeof(MPCConfig.ControllerConfig));
            using (FileStream fs = new FileStream("ItsYourMPCConfig.xml", FileMode.Create/*FileMode.OpenOrCreate*/))
            {
                contrexport.Serialize(fs, controllerConfig);
            }
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
                if (!string.IsNullOrEmpty((string)wMPC.Cells[2, 1].Value))
                {
                    controllerConfig.ActualStateOPCPath = Convert.ToString(wMPC.Cells[2, 1].Value);
                }

                if (!string.IsNullOrEmpty((string)wMPC.Cells[2, 2].Value))
                {
                    controllerConfig.DesiredStateOPCPath = Convert.ToString(wMPC.Cells[2, 2].Value);
                }

                if (!string.IsNullOrEmpty((string)wMPC.Cells[2, 3].Value))
                {
                    controllerConfig.WatchDogOPCPath = Convert.ToString(wMPC.Cells[2, 3].Value);
                }
            }


            for (int i = 1; i < 50; i++)
            {
                for (int j = 1; j < 50; j++)
                {
                    if (!string.IsNullOrEmpty((string)wMatrix.Cells[i, j].Value))
                    {
                        coefs = Convert.ToString(wMatrix.Cells[i, j].Value).Replace(" ", "").Replace("\n", " ").Replace("=", " ").Split(' ', '.');

                        if (coefs.Length == 1 && coefs[0].Contains("CV") && coefs[0].StartsWith('C'))
                        {
                            if (!string.IsNullOrEmpty((string)wMatrix.Cells[i, j + 1].Value))
                            {
                                controllerConfig.CVs.Add(new CVConfig()
                                {
                                    Name = (string)wMatrix.Cells[i, j + 1].Value,
                                    Description = (string)wMatrix.Cells[i, j + 2].Value,
                                    Weigth = 1,
                                    Priority = 1
                                });
                            }


                            try
                            {
                                if (!string.IsNullOrEmpty((string)wCVs.Cells[row_CV, 2].Value))
                                {
                                    controllerConfig.CVs[row_CV - 2].POV = new ComplexDouble()
                                    {
                                        OPCTag = Convert.ToString(wCVs.Cells[row_CV, 2].Value)
                                    };
                                }
                            }
                            catch
                            {
                                if (!double.IsNaN(Convert.ToDouble(wCVs.Cells[row_CV, 2].Value)))
                                {
                                    controllerConfig.CVs[row_CV - 2].POV = new ComplexDouble()
                                    {
                                        Value = Convert.ToDouble(wCVs.Cells[row_CV, 2].Value)
                                    };
                                }
                            }
                            try
                            {
                                if (!string.IsNullOrEmpty((string)wCVs.Cells[row_CV, 3].Value))
                                {
                                    controllerConfig.CVs[row_CV - 2].LoLimitInput = new ComplexDouble()
                                    {
                                        OPCTag = Convert.ToString(wCVs.Cells[row_CV, 3].Value)
                                    };
                                }
                            }
                            catch
                            {
                                if (!double.IsNaN(Convert.ToDouble(wCVs.Cells[row_CV, 3].Value)))
                                {
                                    controllerConfig.CVs[row_CV - 2].LoLimitInput = new ComplexDouble()
                                    {
                                        Value = Convert.ToDouble(wCVs.Cells[row_CV, 3].Value)
                                    };
                                }
                            }
                            try
                            {
                                if (!string.IsNullOrEmpty((string)wCVs.Cells[row_CV, 4].Value))
                                {
                                    controllerConfig.CVs[row_CV - 2].HiLimitInput = new ComplexDouble()
                                    {
                                        OPCTag = Convert.ToString(wCVs.Cells[row_CV, 4].Value)
                                    };
                                }
                            }
                            catch
                            {
                                if (!double.IsNaN(Convert.ToDouble(wCVs.Cells[row_CV, 4].Value)))
                                {
                                    controllerConfig.CVs[row_CV - 2].HiLimitInput = new ComplexDouble()
                                    {
                                        Value = Convert.ToDouble(wCVs.Cells[row_CV, 4].Value)
                                    };
                                }
                            }

                            try
                            {
                                if (!string.IsNullOrEmpty((string)wCVs.Cells[row_CV, 5].Value))
                                {
                                    controllerConfig.CVs[row_CV - 2].LoLimitEng = new ComplexDouble()
                                    {
                                        OPCTag = Convert.ToString(wCVs.Cells[row_CV, 5].Value)
                                    };
                                }
                            }
                            catch
                            {
                                if (!double.IsNaN(Convert.ToDouble(wCVs.Cells[row_CV, 5].Value)))
                                {
                                    controllerConfig.CVs[row_CV - 2].LoLimitEng = new ComplexDouble()
                                    {
                                        Value = Convert.ToDouble(wCVs.Cells[row_CV, 5].Value)
                                    };
                                }
                            }



                            try
                            {
                                if (!string.IsNullOrEmpty((string)wCVs.Cells[row_CV, 6].Value))
                                {
                                    controllerConfig.CVs[row_CV - 2].HiLimitEng = new ComplexDouble()
                                    {
                                        OPCTag = Convert.ToString(wCVs.Cells[row_CV, 6].Value)
                                    };
                                }
                            }
                            catch
                            {
                                if (!double.IsNaN(Convert.ToDouble(wCVs.Cells[row_CV, 6].Value)))
                                {
                                    controllerConfig.CVs[row_CV - 2].HiLimitEng = new ComplexDouble()
                                    {
                                        Value = Convert.ToDouble(wCVs.Cells[row_CV, 6].Value)
                                    };
                                }
                            }

                            try
                            {
                                if (!string.IsNullOrEmpty((string)wCVs.Cells[row_CV, 7].Value))
                                {
                                    controllerConfig.CVs[row_CV - 2].ActualStateOPCPath = Convert.ToString(wCVs.Cells[row_CV, 7].Value);
                                }
                            }
                            catch
                            {

                            }

                            try
                            {
                                if (!string.IsNullOrEmpty((string)wCVs.Cells[row_CV, 8].Value))
                                {
                                    controllerConfig.CVs[row_CV - 2].DesiredStateOPCPath = Convert.ToString(wCVs.Cells[row_CV, 8].Value);
                                }
                            }
                            catch
                            {

                            }


                            try
                            {
                                if (!string.IsNullOrEmpty((string)wCVs.Cells[row_CV, 9].Value))
                                {
                                    controllerConfig.CVs[row_CV - 2].SSValue = new ComplexDouble()
                                    {
                                        OPCTag = Convert.ToString(wCVs.Cells[row_CV, 9].Value)
                                    };
                                }
                            }
                            catch
                            {
                                if (!double.IsNaN(Convert.ToDouble(wCVs.Cells[row_CV, 9].Value)))
                                {
                                    controllerConfig.CVs[row_CV - 2].SSValue = new ComplexDouble()
                                    {
                                        Value = Convert.ToDouble(wCVs.Cells[row_CV, 9].Value)
                                    };
                                }
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
                                    Description = (string)wMatrix.Cells[i + 2, j].Value,
                                    Weigth = 1
                                });
                            }




                            try
                            {
                                if (!string.IsNullOrEmpty((string)wMVs.Cells[row_MV, 2].Value))
                                {
                                    controllerConfig.MVs[row_MV - 2].PV = new ComplexDouble()
                                    {
                                        OPCTag = Convert.ToString(wMVs.Cells[row_MV, 2].Value)
                                    };
                                }
                            }
                            catch
                            {
                                if (!double.IsNaN(Convert.ToDouble(wMVs.Cells[row_MV, 2].Value)))
                                {
                                    controllerConfig.MVs[row_MV - 2].PV = new ComplexDouble()
                                    {
                                        Value = Convert.ToDouble(wMVs.Cells[row_MV, 2].Value)
                                    };
                                }
                            }


                            try
                            {
                                if (!string.IsNullOrEmpty((string)wMVs.Cells[row_MV, 3].Value))
                                {
                                    controllerConfig.MVs[row_MV - 2].SV = new ComplexDouble()
                                    {
                                        OPCTag = Convert.ToString(wMVs.Cells[row_MV, 3].Value)
                                    };
                                }
                            }
                            catch
                            {
                                if (!double.IsNaN(Convert.ToDouble(wMVs.Cells[row_MV, 3].Value)))
                                {
                                    controllerConfig.MVs[row_MV - 2].SV = new ComplexDouble()
                                    {
                                        Value = Convert.ToDouble(wMVs.Cells[row_MV, 3].Value)
                                    };
                                }
                            }


                            try
                            {
                                if (!string.IsNullOrEmpty((string)wMVs.Cells[row_MV, 4].Value))
                                {
                                    controllerConfig.MVs[row_MV - 2].RSV = new ComplexDouble()
                                    {
                                        OPCTag = Convert.ToString(wMVs.Cells[row_MV, 4].Value)
                                    };
                                }
                            }
                            catch
                            {
                                if (!double.IsNaN(Convert.ToDouble(wMVs.Cells[row_MV, 4].Value)))
                                {
                                    controllerConfig.MVs[row_MV - 2].RSV = new ComplexDouble()
                                    {
                                        Value = Convert.ToDouble(wMVs.Cells[row_MV, 4].Value)
                                    };
                                }
                            }


                            try
                            {
                                if (!string.IsNullOrEmpty((string)wMVs.Cells[row_MV, 5].Value))
                                {
                                    controllerConfig.MVs[row_MV - 2].OP = new ComplexDouble()
                                    {
                                        OPCTag = Convert.ToString(wMVs.Cells[row_MV, 5].Value)
                                    };
                                }
                            }
                            catch
                            {
                                if (!double.IsNaN(Convert.ToDouble(wMVs.Cells[row_MV, 5].Value)))
                                {
                                    controllerConfig.MVs[row_MV - 2].OP = new ComplexDouble()
                                    {
                                        Value = Convert.ToDouble(wMVs.Cells[row_MV, 5].Value)
                                    };
                                }
                            }


                            try
                            {
                                if (!double.IsNaN(Convert.ToDouble(wMVs.Cells[row_MV, 6].Value)))
                                {
                                    controllerConfig.MVs[row_MV - 2].dMVup = Convert.ToDouble(wMVs.Cells[row_MV, 6].Value);
                                }
                            }
                            catch
                            {

                            }


                            try
                            {
                                if (!double.IsNaN(Convert.ToDouble(wMVs.Cells[row_MV, 7].Value)))
                                {
                                    controllerConfig.MVs[row_MV - 2].dMVdown = Convert.ToDouble(wMVs.Cells[row_MV, 7].Value);
                                }
                            }
                            catch
                            {

                            }

                            try
                            {
                                if (!string.IsNullOrEmpty((string)wMVs.Cells[row_MV, 8].Value))
                                {
                                    controllerConfig.MVs[row_MV - 2].dMV = new ComplexDouble()
                                    {
                                        OPCTag = Convert.ToString(wMVs.Cells[row_MV, 8].Value)
                                    };
                                }
                            }
                            catch
                            {
                                if (!double.IsNaN(Convert.ToDouble(wMVs.Cells[row_MV, 8].Value)))
                                {
                                    controllerConfig.MVs[row_MV - 2].dMV = new ComplexDouble()
                                    {
                                        Value = Convert.ToDouble(wMVs.Cells[row_MV, 8].Value)
                                    };
                                }
                            }

                            try
                            {
                                if (!string.IsNullOrEmpty((string)wMVs.Cells[row_MV, 9].Value))
                                {
                                    controllerConfig.MVs[row_MV - 2].LoLimitInput = new ComplexDouble()
                                    {
                                        OPCTag = Convert.ToString(wMVs.Cells[row_MV, 9].Value)
                                    };
                                }
                            }
                            catch
                            {
                                if (!double.IsNaN(Convert.ToDouble(wMVs.Cells[row_MV, 9].Value)))
                                {
                                    controllerConfig.MVs[row_MV - 2].LoLimitInput = new ComplexDouble()
                                    {
                                        Value = Convert.ToDouble(wMVs.Cells[row_MV, 9].Value)
                                    };
                                }
                            }

                            try
                            {
                                if (!string.IsNullOrEmpty((string)wMVs.Cells[row_MV, 10].Value))
                                {
                                    controllerConfig.MVs[row_MV - 2].HiLimitInput = new ComplexDouble()
                                    {
                                        OPCTag = Convert.ToString(wMVs.Cells[row_MV, 10].Value)
                                    };
                                }
                            }
                            catch
                            {
                                if (!double.IsNaN(Convert.ToDouble(wMVs.Cells[row_MV, 10].Value)))
                                {
                                    controllerConfig.MVs[row_MV - 2].HiLimitInput = new ComplexDouble()
                                    {
                                        Value = Convert.ToDouble(wMVs.Cells[row_MV, 10].Value)
                                    };
                                }
                            }


                            try
                            {
                                if (!string.IsNullOrEmpty((string)wMVs.Cells[row_MV, 11].Value))
                                {
                                    controllerConfig.MVs[row_MV - 2].LoLimitEng = new ComplexDouble()
                                    {
                                        OPCTag = Convert.ToString(wMVs.Cells[row_MV, 11].Value)
                                    };
                                }
                            }
                            catch
                            {
                                if (!double.IsNaN(Convert.ToDouble(wMVs.Cells[row_MV, 11].Value)))
                                {
                                    controllerConfig.MVs[row_MV - 2].LoLimitEng = new ComplexDouble()
                                    {
                                        Value = Convert.ToDouble(wMVs.Cells[row_MV, 11].Value)
                                    };
                                }
                            }






                            try
                            {
                                if (!string.IsNullOrEmpty((string)wMVs.Cells[row_MV, 12].Value))
                                {
                                    controllerConfig.MVs[row_MV - 2].HiLimitEng = new ComplexDouble()
                                    {
                                        OPCTag = Convert.ToString(wMVs.Cells[row_MV, 12].Value)
                                    };
                                }
                            }
                            catch
                            {
                                if (!double.IsNaN(Convert.ToDouble(wMVs.Cells[row_MV, 12].Value)))
                                {
                                    controllerConfig.MVs[row_MV - 2].HiLimitEng = new ComplexDouble()
                                    {
                                        Value = Convert.ToDouble(wMVs.Cells[row_MV, 12].Value)
                                    };
                                }
                            }

                            try
                            {
                                if (!string.IsNullOrEmpty((string)wMVs.Cells[row_MV, 13].Value))
                                {
                                    controllerConfig.MVs[row_MV - 2].ActualStateOPCPath = Convert.ToString(wMVs.Cells[row_MV, 13].Value);

                                }
                            }
                            catch
                            {

                            }

                            try
                            {
                                if (!string.IsNullOrEmpty((string)wMVs.Cells[row_MV, 14].Value))
                                {
                                    controllerConfig.MVs[row_MV - 2].DesiredStateOPCPath = Convert.ToString(wMVs.Cells[row_MV, 14].Value);

                                }
                            }
                            catch
                            {

                            }


                            try
                            {
                                if (!string.IsNullOrEmpty((string)wMVs.Cells[row_MV, 15].Value))
                                {
                                    controllerConfig.MVs[row_MV - 2].SSValue = new ComplexDouble()
                                    {
                                        OPCTag = Convert.ToString(wMVs.Cells[row_MV, 15].Value)
                                    };
                                }
                            }
                            catch
                            {
                                if (!double.IsNaN(Convert.ToDouble(wMVs.Cells[row_MV, 15].Value)))
                                {
                                    controllerConfig.MVs[row_MV - 2].SSValue = new ComplexDouble()
                                    {
                                        Value = Convert.ToDouble(wMVs.Cells[row_MV, 15].Value)
                                    };
                                }
                            }


                            try
                            {
                                if (!string.IsNullOrEmpty((string)wMVs.Cells[row_MV, 16].Value))
                                {
                                    controllerConfig.MVs[row_MV - 2].CalculatedValue = new ComplexDouble()
                                    {
                                        OPCTag = Convert.ToString(wMVs.Cells[row_MV, 16].Value)
                                    };
                                }
                            }
                            catch
                            {
                                if (!double.IsNaN(Convert.ToDouble(wMVs.Cells[row_MV, 16].Value)))
                                {
                                    controllerConfig.MVs[row_MV - 2].CalculatedValue = new ComplexDouble()
                                    {
                                        Value = Convert.ToDouble(wMVs.Cells[row_MV, 16].Value)
                                    };
                                }
                            }

                            try
                            {
                                if (!string.IsNullOrEmpty((string)wMVs.Cells[row_MV, 17].Value))
                                {
                                    controllerConfig.MVs[row_MV - 2].OPHi = new ComplexDouble()
                                    {
                                        OPCTag = Convert.ToString(wMVs.Cells[row_MV, 17].Value)
                                    };
                                }
                            }
                            catch
                            {
                                if (!double.IsNaN(Convert.ToDouble(wMVs.Cells[row_MV, 17].Value)))
                                {
                                    controllerConfig.MVs[row_MV - 2].OPHi = new ComplexDouble()
                                    {
                                        Value = Convert.ToDouble(wMVs.Cells[row_MV, 17].Value)
                                    };
                                }
                            }

                            try
                            {
                                if (!string.IsNullOrEmpty((string)wMVs.Cells[row_MV, 18].Value))
                                {
                                    controllerConfig.MVs[row_MV - 2].OPLo = new ComplexDouble()
                                    {
                                        OPCTag = Convert.ToString(wMVs.Cells[row_MV, 18].Value)
                                    };
                                }
                            }
                            catch
                            {
                                if (!double.IsNaN(Convert.ToDouble(wMVs.Cells[row_MV, 18].Value)))
                                {
                                    controllerConfig.MVs[row_MV - 2].OPLo = new ComplexDouble()
                                    {
                                        Value = Convert.ToDouble(wMVs.Cells[row_MV, 18].Value)
                                    };
                                }
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
                                    Description = (string)wMatrix.Cells[i + 2, j].Value
                                });
                            }



                            try
                            {
                                if (!string.IsNullOrEmpty((string)wDVs.Cells[row_DV, 2].Value))
                                {
                                    controllerConfig.DVs[row_DV - 2].Value = new ComplexDouble()
                                    {
                                        OPCTag = Convert.ToString(wDVs.Cells[row_DV, 2].Value)
                                    };
                                }
                            }
                            catch
                            {
                                if (!double.IsNaN(Convert.ToDouble(wDVs.Cells[row_DV, 2].Value)))
                                {
                                    controllerConfig.DVs[row_DV - 2].Value = new ComplexDouble()
                                    {
                                        Value = Convert.ToDouble(wDVs.Cells[row_DV, 2].Value)
                                    };
                                }
                            }


                            try
                            {
                                if (!string.IsNullOrEmpty((string)wDVs.Cells[row_DV, 3].Value))
                                {
                                    controllerConfig.DVs[row_DV - 2].ActualStateOPCPath = Convert.ToString(wDVs.Cells[row_DV, 3].Value);

                                }
                            }
                            catch
                            {

                            }

                            try
                            {
                                if (!string.IsNullOrEmpty((string)wDVs.Cells[row_DV, 4].Value))
                                {
                                    controllerConfig.DVs[row_DV - 2].DesiredStateOPCPath = Convert.ToString(wDVs.Cells[row_DV, 4].Value);

                                }
                            }
                            catch
                            {

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
                                controllerConfig.Models.Add(new ModelConfig()
                                {
                                    Gain = Convert.ToDouble(coefs[1]),

                                    T = Convert.ToDouble(coefs[5]),

                                    tau = Convert.ToDouble(coefs[3]),

                                    cvindex = i - 4,

                                    mvindex = j - 4,

                                    TypePF = TypesPF.FirstOrder
                                });
                            }
                            catch
                            {
                                continue;
                            }

                        }
                        else if (coefs.Contains("G") && coefs.Contains("D") && !coefs.Contains("T") && !coefs.Contains("T2"))
                        {
                            try
                            {
                                controllerConfig.Models.Add(new ModelConfig()
                                {
                                    Gain = Convert.ToDouble(coefs[1]),

                                    tau = Convert.ToDouble(coefs[3]),

                                    isIntegrity = true,

                                    cvindex = i - 4,

                                    mvindex = j - 4,

                                    TypePF = TypesPF.Integer
                                });
                            }
                            catch
                            {
                                continue;
                            }

                        }
                        else if (coefs.Contains("G") && coefs.Contains("D") && coefs.Contains("T") && coefs.Contains("T2"))
                        {
                            try
                            {
                                controllerConfig.Models.Add(new ModelConfig()
                                {
                                    Gain = Convert.ToDouble(coefs[1]),

                                    T = Convert.ToDouble(coefs[5]),

                                    T2 = Convert.ToDouble(coefs[7]),

                                    tau = Convert.ToDouble(coefs[3]),

                                    cvindex = i - 4,

                                    mvindex = j - 4,

                                    TypePF = TypesPF.SecondOrder
                                });
                            }
                            catch
                            {
                                continue;
                            }

                        }


                    }
                }
            }
        }


        class ReadModel
        {
            string _NameMV;
            string _NameCV;
            double _Gain;

            public string NameMV { get => _NameMV; set => _NameMV = value; }
            public string NameCV { get => _NameCV; set => _NameCV = value; }
            public double Gain { get => _Gain; set => _Gain = value; }
        }

        class Model
        {
            public int mvindex { get; set; }
            public int cvindex { get; set; }

            public double Gain { get; set; }
        }

        class CV
        {
            public double Value { get; set; }
            public string Name { get; set; }
        }

        class MV
        {
            public double Value { get; set; }
            public string Name { get; set; }
        }
    }
}
