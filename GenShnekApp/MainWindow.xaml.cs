using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Kompas6API5;
using Kompas6Constants3D;
using Kompas6Constants;
using System.Runtime.InteropServices;
using KompasAPI7;
using System.IO;
using static System.Math;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Win32;

namespace GenShnekApp
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        double holeDiam;
        double tubeLength;
        double shnekThick;
        double shnekDiam;
        double hexSize;
        double hex2Size;
        double holeDistance;
        double tubeRad;
        double step;

        double type2ShnekDiam;
        double type2T;
        double type2T1;
        double type2T2;
        double threadDiam;
        double threadStep;

        double extrDiam;
        double extrRad;
        double extrCoffLength;
        double extrLength;

        string extrMethod = "—";
        string extrName = "—";

        KompasObject kompas;
        ksPart part;

        int typeCount;
        int styleCount;

        bool mistakeCheck;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void GhostTypeSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //Выбор ГОСТа (строго/произвольно)
            ShnekType.Items.Clear();
            switch (GhostType.SelectedIndex)
            {
                case 0:
                    GOSTSelection1();
                    break;
                case 1:
                    GOSTSelection2();
                    break;
                case 2:
                    DefaultShnekChoose.Items.Clear();
                    GOSTSelection3();
                    break;
                case 3:
                    GOSTSelection4();
                    break;
            }
        }

        //Выбор Типа шнека
        private void ShnekTypeSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (GhostType.SelectedIndex == 0 || GhostType.SelectedIndex == 1)
            {
                ShnekStyle.Items.Clear();
                DefaultShnekChoose.Items.Clear();
                if (ShnekType.SelectedIndex == 0 || ShnekType.SelectedIndex == 1)
                {
                    switch (ShnekType.SelectedIndex)
                    {
                        case 0:
                            if (ImgSketch != null) ImgSketch.Source = (ImageSource)new ImageSourceConverter().ConvertFrom(new Uri(@"D:\Users\Garnik\Desktop\учёба\Диплом\GenShnekApp\GenShnekApp\ShnekSketch1.png"));
                            if (ImgTable != null) ImgTable.Source = (ImageSource)new ImageSourceConverter().ConvertFrom(new Uri(@"D:\Users\Garnik\Desktop\учёба\Диплом\GenShnekApp\GenShnekApp\ShnekTable1.png"));
                            styleCount = 2;
                            if (GhostType.SelectedIndex == 0) ShnekStyle.IsEnabled = false;
                            DefaultShnekItems1();
                            break;
                        case 1:
                            if (ImgTable != null) ImgTable.Source = (ImageSource)new ImageSourceConverter().ConvertFrom(new Uri(@"D:\Users\Garnik\Desktop\учёба\Диплом\GenShnekApp\GenShnekApp\ShnekTable2.png"));
                            styleCount = 2;
                            DefaultShnekItems2();
                            ShnekStyle.IsEnabled = true;
                            break;
                    }
                }
                for (int i = 0; i < styleCount; i++) ShnekStyle.Items.Add($"Исполнение {i + 1}");
                ShnekStyle.SelectedIndex = 0;
            }
        }

        private void ShnekStyleSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (GhostType.SelectedIndex == 0 || GhostType.SelectedIndex == 1)
            {
                if (ShnekType.SelectedIndex == 1)
                {
                    if (ShnekStyle.SelectedIndex == 0) if (ImgSketch != null) ImgSketch.Source = (ImageSource)new ImageSourceConverter().ConvertFrom(new Uri(@"D:\Users\Garnik\Desktop\учёба\Диплом\GenShnekApp\GenShnekApp\ShnekSketch21.png"));
                    if (ShnekStyle.SelectedIndex == 1) if (ImgSketch != null) ImgSketch.Source = (ImageSource)new ImageSourceConverter().ConvertFrom(new Uri(@"D:\Users\Garnik\Desktop\учёба\Диплом\GenShnekApp\GenShnekApp\ShnekSketch22.png"));
                }
            }
            if (GhostType.SelectedIndex == 1)
            {
                if (ShnekType.SelectedIndex == 0)
                {
                    if (ShnekStyle.SelectedIndex == 0)
                    {
                        inputHexSize.IsEnabled = true;
                        inputHex2Size.IsEnabled = false;
                    }
                    if (ShnekStyle.SelectedIndex == 1)
                    {
                        inputHexSize.IsEnabled = false;
                        inputHex2Size.IsEnabled = true;
                    }
                }
                else if (ShnekType.SelectedIndex ==1)
                {
                    if (ShnekStyle.SelectedIndex == 0)
                    {
                        inputType2T.IsEnabled = false;
                        inputType2T1.IsEnabled = true;
                        inputType2T2.IsEnabled = true;
                    }
                    if (ShnekStyle.SelectedIndex == 1)
                    {
                        inputType2T.IsEnabled = true;
                        inputType2T1.IsEnabled = false;
                        inputType2T2.IsEnabled = false;
                    }
                }
            }
        }

        private void TextBoxInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "0123456789,".IndexOf(e.Text) < 0;
        }
        private void DeleteSpaces(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Space)
            {
                e.Handled = true;
            }
        }

        private void CreationButton(object sender, RoutedEventArgs e)
        {
            ParamConv();
            if (mistakeCheck == true)
            {
                e.Handled = false;
            }
            if (mistakeCheck == false)
            {
                e.Handled = true;
                return;
            }

            // Удалить когда модуль станет интегрированным
            try
            {
                kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5");
            }
            catch
            {
                kompas = (KompasObject)Activator.CreateInstance(Type.GetTypeFromProgID("KOMPAS.Application.5"));
            }
            //kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5"); //использовать когда модуль станет интегрированным
            if (kompas == null) return;
            kompas.Visible = true;

            ksDocument3D ksDoc3d = kompas.Document3D(); // создание 3д документа

            ksDoc3d.Create(false, true); // false - видимый режим, true - деталь
            ksDoc3d = kompas.ActiveDocument3D(); // указатель на интерфейс 3д модели 
            ksDoc3d.author = "Garnik";   // указание имени автора

            part = ksDoc3d.GetPart((int)Part_Type.pTop_Part); // новый компонент

            //Буровые шнеки
            if (GhostType.SelectedIndex == 0 || GhostType.SelectedIndex == 1)
            {
                //Шнеки первого типа
                if (ShnekType.SelectedIndex == 0)
                {
                    tubeRad = hexSize * 0.75;
                    //Готовые шнеки
                    if (DefaultShnekChoose.IsEnabled == true)
                    {
                        switch (DefaultShnekChoose.SelectedIndex)
                        {
                            case 0:
                                CylinderCreation(55 * 0.75, tubeLength);
                                JointCreation1(55, 52);
                                JointHoleCreation(24, 52, 0);
                                JointHoleCreation(24, -tubeLength + (52 * 2 / 3), 0);
                                SpyralCreation(55 * 0.75, 100, 0, tubeLength, 2, 135);
                                break;
                            case 1:
                                CylinderCreation(55 * 0.75, tubeLength);
                                JointCreation1(55, 52);
                                JointHoleCreation(24, 52, 0);
                                JointHoleCreation(24, -tubeLength + (52 * 2 / 3), 0);
                                SpyralCreation(55 * 0.75, 100, 0, tubeLength, 2, 150);
                                break;
                            case 2:
                                CylinderCreation(55 * 0.75, tubeLength);
                                JointCreation1(55, 52);
                                JointHoleCreation(24, 52, 0);
                                JointHoleCreation(24, -tubeLength + (52 * 2 / 3), 0);
                                SpyralCreation(55 * 0.75, 100, 0, tubeLength, 2, 180);
                                break;
                            case 3:
                                CylinderCreation(60 * 0.75, tubeLength);
                                JointCreation1(60, 55);
                                JointHoleCreation(27, 55, 0);
                                JointHoleCreation(24, -tubeLength + (55 * 2 / 3), 0);
                                SpyralCreation(60 * 0.75, 100, 0, tubeLength, 2, 200);
                                break;
                            case 4:
                                CylinderCreation(60 * 0.75, tubeLength);
                                JointCreation1(60, 55);
                                JointHoleCreation(27, 55, 0);
                                JointHoleCreation(24, -tubeLength + (55 * 2 / 3), 0);
                                SpyralCreation(60 * 0.75, 100, 0, tubeLength, 2, 300);
                                break;
                            case 5:
                                JointCreation2(90, 95 * 3 / 2);
                                CylinderCreation(90 * 0.75, tubeLength);
                                JointHoleCreation(30, 95, 0);
                                JointHoleCreation(24, -tubeLength + (95 * 2 / 3), 0);
                                SpyralCreation(90 * 0.75, 100, 0, tubeLength, 2, 300);
                                break;
                        }
                    }
                    //Выбор исполнения шнека
                    if (ShnekStyle.IsEnabled == true)
                    {
                        switch (ShnekStyle.SelectedIndex)
                        {
                            case 0:
                                CylinderCreation(tubeRad, tubeLength);
                                JointCreation1(hexSize, holeDistance);
                                JointHoleCreation(holeDiam, holeDistance, 0);
                                JointHoleCreation(holeDiam, -tubeLength + (holeDistance * 2 / 3), 0);
                                SpyralCreation(tubeRad, step, 0, tubeLength, shnekThick, shnekDiam);
                                break;
                            case 1:
                                JointCreation2(hex2Size, holeDistance * 3 / 2);
                                CylinderCreation(tubeRad, tubeLength);
                                JointHoleCreation(holeDiam, holeDistance, 0);
                                JointHoleCreation(holeDiam, -tubeLength + (holeDistance * 2 / 3), 0);
                                SpyralCreation(tubeRad, step, 0, tubeLength, shnekThick, shnekDiam);
                                break;
                        }
                    }
                }
                //Шнеки второго типа
                else
                {
                    //Готовые шнеки
                    if (DefaultShnekChoose.IsEnabled == true)
                    {
                        switch (DefaultShnekChoose.SelectedIndex)
                        {
                            case 0:
                                type2ShnekDiam = 80;
                                tubeRad = (type2ShnekDiam * 10 / 18) / 2;
                                CylinderCreation(tubeRad, tubeLength);
                                SpyralCreation(tubeRad, step, 0, tubeLength, shnekThick, type2ShnekDiam);
                                JointCreation4(32, 28, 56);
                                HoleType2Creation2(32, 56, tubeLength);
                                SpyralCreation(28 / 2, 6, tubeLength - (56 * 7 / 8), 56 * 3 / 4, 3, 32);
                                SpyralCreation(28 / 2, 6, -56 * 4 / 3 * 0.95, 56, 3, 32);
                                break;
                            case 1:
                                type2ShnekDiam = 100;
                                tubeRad = (type2ShnekDiam * 10 / 18) / 2;
                                CylinderCreation(tubeRad, tubeLength);
                                SpyralCreation(tubeRad, step, 0, tubeLength, shnekThick, type2ShnekDiam);
                                JointCreation4(40, 36, 63);
                                HoleType2Creation2(40 , 63, tubeLength);
                                SpyralCreation(36 / 2, 8, tubeLength - (63 * 7 / 8), 63 * 3 / 4, 4, 40);
                                SpyralCreation(36 / 2, 8, -63 * 4 / 3 * 0.95, 63, 4, 40);
                                break;
                            case 2:
                                type2ShnekDiam = 200;
                                tubeRad = (type2ShnekDiam * 10 / 18) / 2;
                                CylinderCreation(tubeRad, tubeLength);
                                SpyralCreation(tubeRad, step, 0, tubeLength, shnekThick, type2ShnekDiam);
                                JointCreation3(83, tubeLength, 324/2);
                                HoleType2Creation1(95, 163, 324, tubeLength);
                                SpyralCreation(41.5, 16, 163 / 16, 163 * 3 / 4, 7.75, 95);
                                SpyralCreation(41.5, 16, tubeLength + 324 / 16, 324 * 3 / 8, 8.1, 95);
                                break;
                        }
                    }
                    //Выбор исполнения шнека
                    else
                    {
                        double threadSemiStep = threadStep / 2;
                        double threadDiam0 = threadDiam * 0.87;
                        double threadRad0 = threadDiam0 / 2;
                        tubeRad = (type2ShnekDiam * 10 / 18) / 2;
                        if (ShnekStyle.SelectedIndex == 0)
                        {
                            CylinderCreation(tubeRad, tubeLength);
                            SpyralCreation(tubeRad, step, 0, tubeLength, shnekThick, type2ShnekDiam);
                            JointCreation3(threadDiam0, tubeLength, type2T2 / 2);
                            HoleType2Creation1(threadDiam, type2T1, type2T2, tubeLength);
                            SpyralCreation(threadRad0, threadStep, type2T1 / 16, type2T1 * 3 / 4, threadSemiStep, threadDiam);
                            SpyralCreation(threadRad0, threadStep, tubeLength + type2T2 / 16, type2T2 * 3 / 8, threadSemiStep, threadDiam);
                        }
                        else
                        {
                            CylinderCreation(tubeRad, tubeLength);
                            SpyralCreation(tubeRad, step, 0, tubeLength, shnekThick, type2ShnekDiam);
                            JointCreation4(threadDiam, threadDiam0, type2T);
                            HoleType2Creation2(threadDiam, type2T, tubeLength);
                            SpyralCreation(threadRad0, threadStep, tubeLength - (type2T * 7 / 8), type2T * 3 / 4, threadSemiStep, threadDiam);
                            SpyralCreation(threadRad0, threadStep, -type2T * 4 / 3 * 0.95, type2T, threadSemiStep, threadDiam);
                        }
                    }
                }
            }
            //Экструзионные шнеки
            else
            {
                extrRad = extrDiam / 2;
                if (DefaultShnekChoose.IsEnabled == true)
                {
                    //Готовые шнеки
                    extrMethod = "стандартный";
                    switch (DefaultShnekChoose.SelectedIndex)
                    {
                        case 0:
                            extrDiam = 20;
                            extrCoffLength = 20;
                            extrLength = extrDiam * extrCoffLength;
                            break;
                        case 1:
                            extrDiam = 32;
                            extrCoffLength = 20;
                            extrLength = extrDiam * extrCoffLength;
                            break;
                        case 2:
                            extrDiam = 45;
                            extrCoffLength = 20;
                            extrLength = extrDiam * extrCoffLength;
                            break;
                        case 3:
                            extrDiam = 45;
                            extrCoffLength = 25;
                            extrLength = extrDiam * extrCoffLength;
                            break;
                        case 4:
                            extrDiam = 63;
                            extrCoffLength = 20;
                            extrLength = extrDiam * extrCoffLength;
                            break;
                        case 5:
                            extrDiam = 63;
                            extrCoffLength = 25;
                            extrLength = extrDiam * extrCoffLength;
                            break;
                        case 6:
                            extrDiam = 63;
                            extrCoffLength = 30;
                            extrLength = extrDiam * extrCoffLength;
                            break;
                        case 7:
                            extrDiam = 90;
                            extrCoffLength = 20;
                            extrLength = extrDiam * extrCoffLength;
                            break;
                        case 8:
                            extrDiam = 90;
                            extrCoffLength = 25;
                            extrLength = extrDiam * extrCoffLength;
                            break;
                        case 9:
                            extrDiam = 90;
                            extrCoffLength = 30;
                            extrLength = extrDiam * extrCoffLength;
                            break;
                        case 10:
                            extrDiam = 125;
                            extrCoffLength = 25;
                            extrLength = extrDiam * extrCoffLength;
                            break;
                        case 11:
                            extrDiam = 160;
                            extrCoffLength = 20;
                            extrLength = extrDiam * extrCoffLength;
                            break;
                        case 12:
                            extrDiam = 200;
                            extrCoffLength = 20;
                            extrLength = extrDiam * extrCoffLength;
                            break;
                    }
                    extrName = $"ЧП {extrDiam}x{extrCoffLength}";
                    extrRad = extrDiam / 2;
                    CylinderCreation(extrRad, extrLength);
                    SpyralCreation(extrRad * 1.2, extrDiam * 1.2, 0, extrLength, extrDiam * 0.06, extrDiam);
                    ConeCreation(extrRad);
                    ShnekCalc(extrDiam, extrLength);
                }
                //Пользовательский шнек
                else
                {
                    extrMethod = "пользовательский";
                    extrName = "отсутствует";
                    extrLength = extrDiam * extrCoffLength;
                    CylinderCreation(extrRad, extrLength);
                    SpyralCreation(extrRad * 1.2, extrDiam * 1.2, 0, extrLength, extrDiam * 0.06, extrDiam);
                    ConeCreation(extrRad);
                    ShnekCalc(extrDiam, extrLength);
                }
            }
        }

        ///////////////////////////Создание трубы шнека/////////////////////////////
        private void CylinderCreation(double rad, double length)
        {
            ksEntity basePlaneZOY = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeYOZ);

            ksEntity ksSketchE = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef = ksSketchE.GetDefinition();

            ksSketchDef.SetPlane(basePlaneZOY);
            ksSketchE.Create();
            ksDocument2D Sketch2D = (ksDocument2D)ksSketchDef.BeginEdit();

            Sketch2D.ksCircle(0, 0, rad, 1);

            ksSketchDef.EndEdit();

            ksEntity baseExtr = part.NewEntity((short)Obj3dType.o3d_baseExtrusion);
            ksBaseExtrusionDefinition extrDef = baseExtr.GetDefinition();
            ksExtrusionParam extrProp = (ksExtrusionParam)extrDef.ExtrusionParam();

            if (extrProp != null)
            {
                extrDef.SetSketch(ksSketchE);

                extrProp.direction = (short)Direction_Type.dtNormal;
                extrProp.typeNormal = (short)End_Type.etBlind;
                extrProp.depthNormal = length;
                baseExtr.Create();
            }
        }

        ///////////////////////////Создание отверстия шестигранника/////////////////////////////
        private void JointHoleCreation(double diam, double x, double y)
        {
            ksEntity basePlaneXOZ = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOZ);

            ksEntity ksSketchE = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef = ksSketchE.GetDefinition();

            ksSketchDef.SetPlane(basePlaneXOZ);
            ksSketchE.Create();
            ksDocument2D Sketch2D = (ksDocument2D)ksSketchDef.BeginEdit();

            Sketch2D.ksCircle(x, y, diam / 2, 1);

            ksSketchDef.EndEdit();

            ksEntity cutExtr = part.NewEntity((short)Obj3dType.o3d_cutExtrusion);
            ksCutExtrusionDefinition cutDef = cutExtr.GetDefinition();
            ksExtrusionParam cutProp = (ksExtrusionParam)cutDef.ExtrusionParam();

            if (cutProp != null)
            {
                cutDef.SetSketch(ksSketchE);

                cutProp.direction = (short)Direction_Type.dtBoth;
                cutProp.typeNormal = (short)End_Type.etThroughAll;
                cutProp.typeReverse = (short)End_Type.etThroughAll;
                cutExtr.Create();
            }
        }

        ///////////////////////////Создание присоединительного элемента 1 (тип 1 исполнение 1)/////////////////////////////
        private void JointCreation1(double size, double length)
        {
            ksEntity basePlaneZOY = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeYOZ);

            ksEntity ksSketchE = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef = ksSketchE.GetDefinition();

            ksSketchDef.SetPlane(basePlaneZOY);
            ksSketchE.Create();
            ksDocument2D Sketch2D = (ksDocument2D)ksSketchDef.BeginEdit();

            ksRegularPolygonParam hex = (ksRegularPolygonParam)kompas.GetParamStruct((short)StructType2DEnum.ko_RegularPolygonParam);

            if (hex != null)
            {
                hex.xc = 0;
                hex.yc = 0;
                hex.ang = 90;
                hex.count = 6;
                hex.describe = true;
                hex.radius = size / 2;
                hex.style = 1;
                Sketch2D.ksRegularPolygon(hex);
            }

            ksSketchDef.EndEdit();

            ksEntity baseExtr = part.NewEntity((short)Obj3dType.o3d_baseExtrusion);
            ksBaseExtrusionDefinition extrDef = baseExtr.GetDefinition();
            ksExtrusionParam extrProp = (ksExtrusionParam)extrDef.ExtrusionParam();

            if (extrProp != null)
            {
                extrDef.SetSketch(ksSketchE);

                extrProp.direction = (short)Direction_Type.dtReverse;
                extrProp.typeNormal = (short)End_Type.etBlind;
                extrProp.depthReverse = length * 3 / 2;
                baseExtr.Create();
            }
        }

        ///////////////////////////Создание присоединительного элемента 2 (тип 1 исполнение 2)/////////////////////////////
        private void JointCreation2(double diam, double length)
        {
            ksEntity basePlaneZOY = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeYOZ);

            ksEntity ksSketchE1 = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef1 = ksSketchE1.GetDefinition();

            ksSketchDef1.SetPlane(basePlaneZOY);
            ksSketchE1.Create();
            ksDocument2D Sketch2D1 = (ksDocument2D)ksSketchDef1.BeginEdit();

            Sketch2D1.ksCircle(0, 0, diam / 2, 1);

            ksSketchDef1.EndEdit();

            ksEntity bossExtr1 = part.NewEntity((short)Obj3dType.o3d_baseExtrusion);
            ksBaseExtrusionDefinition extrDef1 = bossExtr1.GetDefinition();
            ksExtrusionParam extrProp1 = (ksExtrusionParam)extrDef1.ExtrusionParam();

            if (extrProp1 != null)
            {
                extrDef1.SetSketch(ksSketchE1);

                extrProp1.direction = (short)Direction_Type.dtReverse;
                extrProp1.typeNormal = (short)End_Type.etBlind;
                extrProp1.depthReverse = length;
                bossExtr1.Create();
            }

            double size = diam * 80 / 90;

            ksEntity ksSketchE2 = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef2 = ksSketchE2.GetDefinition();

            ksSketchDef2.SetPlane(basePlaneZOY);
            ksSketchE2.Create();
            ksDocument2D Sketch2D2 = (ksDocument2D)ksSketchDef2.BeginEdit();

            ksRegularPolygonParam triangle = (ksRegularPolygonParam)kompas.GetParamStruct((short)StructType2DEnum.ko_RegularPolygonParam);

            if (triangle != null)
            {
                triangle.xc = 0;
                triangle.yc = 0;
                triangle.ang = 270;
                triangle.count = 3;
                triangle.describe = true;
                triangle.radius = size / 2;
                triangle.style = 1;
                Sketch2D2.ksRegularPolygon(triangle);
            }

            ksSketchDef2.EndEdit();

            ksEntity bossExtr2 = part.NewEntity((short)Obj3dType.o3d_cutExtrusion);
            ksCutExtrusionDefinition extrDef2 = bossExtr2.GetDefinition();
            ksExtrusionParam extrProp2 = (ksExtrusionParam)extrDef2.ExtrusionParam();

            if (extrProp2 != null)
            {
                extrDef2.SetSketch(ksSketchE2);
                extrDef2.cut = false;

                extrProp2.direction = (short)Direction_Type.dtNormal;
                extrProp2.typeNormal = (short)End_Type.etBlind;
                extrProp2.depthNormal = length;
                bossExtr2.Create();
            }

        }

        ///////////////////////////Создание присоединительного элемента 3 (тип 2 исполнение 1)/////////////////////////////
        private void JointCreation3(double threadMinDiam, double length, double jointLength)
        {
            ksEntity basePlaneZOY = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeYOZ);

            double rad1 = threadMinDiam / 2;
            double len1 = jointLength * 0.1;
            double len2 = jointLength * 0.8;
            double len3 = jointLength * 0.05;
            double len4 = jointLength * 0.05;

            ksEntity plane1 = OffsetPlaneCreation(length, basePlaneZOY);
            ksEntity ksSketchE1 = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef1 = ksSketchE1.GetDefinition();

            ksSketchDef1.SetPlane(plane1);
            ksSketchE1.Create();
            ksDocument2D Sketch2D1 = (ksDocument2D)ksSketchDef1.BeginEdit();

            Sketch2D1.ksCircle(0, 0, rad1, 1);

            ksSketchDef1.EndEdit();

            ksEntity bossExtr1 = part.NewEntity((short)Obj3dType.o3d_baseExtrusion);
            ksBaseExtrusionDefinition extrDef1 = bossExtr1.GetDefinition();
            ksExtrusionParam extrProp1 = (ksExtrusionParam)extrDef1.ExtrusionParam();

            if (extrProp1 != null)
            {
                extrDef1.SetSketch(ksSketchE1);

                extrProp1.direction = (short)Direction_Type.dtNormal;
                extrProp1.typeNormal = (short)End_Type.etBlind;
                extrProp1.depthNormal = len1;
                bossExtr1.Create();
            }


            ksEntity plane2 = OffsetPlaneCreation(length + len1, basePlaneZOY);
            ksEntity ksSketchE2 = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef2 = ksSketchE2.GetDefinition();

            ksSketchDef2.SetPlane(plane2);
            ksSketchE2.Create();
            ksDocument2D Sketch2D2 = (ksDocument2D)ksSketchDef2.BeginEdit();

            Sketch2D2.ksCircle(0, 0, rad1, 1);

            ksSketchDef2.EndEdit();

            ksEntity bossExtr2 = part.NewEntity((short)Obj3dType.o3d_baseExtrusion);
            ksBaseExtrusionDefinition extrDef2 = bossExtr2.GetDefinition();
            ksExtrusionParam extrProp2 = (ksExtrusionParam)extrDef2.ExtrusionParam();

            if (extrProp2 != null)
            {
                extrDef2.SetSketch(ksSketchE2);

                extrProp2.direction = (short)Direction_Type.dtNormal;
                extrProp2.typeNormal = (short)End_Type.etBlind;
                extrProp2.depthNormal = len2;
                bossExtr2.Create();
            }

            ksEntity plane3 = OffsetPlaneCreation(length + len1 + len2, basePlaneZOY);
            ksEntity ksSketchE3 = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef3 = ksSketchE3.GetDefinition();

            ksSketchDef3.SetPlane(plane3);
            ksSketchE3.Create();
            ksDocument2D Sketch2D3 = (ksDocument2D)ksSketchDef3.BeginEdit();

            Sketch2D3.ksCircle(0, 0, rad1, 1);

            ksSketchDef3.EndEdit();

            ksEntity bossExtr3 = part.NewEntity((short)Obj3dType.o3d_baseExtrusion);
            ksBaseExtrusionDefinition extrDef3 = bossExtr3.GetDefinition();
            ksExtrusionParam extrProp3 = (ksExtrusionParam)extrDef3.ExtrusionParam();

            if (extrProp3 != null)
            {
                extrDef3.SetSketch(ksSketchE3);

                extrProp3.direction = (short)Direction_Type.dtNormal;
                extrProp3.typeNormal = (short)End_Type.etBlind;
                extrProp3.depthNormal = len3;
                bossExtr3.Create();
            }


            ksEntity plane4 = OffsetPlaneCreation(length + len1 + len2 + len3, basePlaneZOY);
            ksEntity ksSketchE4 = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef4 = ksSketchE4.GetDefinition();

            ksSketchDef4.SetPlane(plane4);
            ksSketchE4.Create();
            ksDocument2D Sketch2D4 = (ksDocument2D)ksSketchDef4.BeginEdit();

            Sketch2D4.ksCircle(0, 0, rad1, 1);

            ksSketchDef4.EndEdit();

            ksEntity bossExtr4 = part.NewEntity((short)Obj3dType.o3d_baseExtrusion);
            ksBaseExtrusionDefinition extrDef4 = bossExtr4.GetDefinition();
            ksExtrusionParam extrProp4 = (ksExtrusionParam)extrDef4.ExtrusionParam();

            if (extrProp4 != null)
            {
                extrDef4.SetSketch(ksSketchE4);

                extrProp4.direction = (short)Direction_Type.dtNormal;
                extrProp4.typeNormal = (short)End_Type.etBlind;
                extrProp4.depthNormal = len4;
                extrProp4.draftOutwardNormal = true;
                extrProp4.draftValueNormal = 45;
                bossExtr4.Create();
            }
        }

        ///////////////////////////Создание присоединительного элемента 4 (тип 2 исполнение 2)/////////////////////////////
        private void JointCreation4(double threadMaxDiam, double threadMinDiam, double length)
        {
            ksEntity basePlaneZOY = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeYOZ);

            length = length * 4 / 3;
            double rad1 = threadMaxDiam / 2;
            double rad2 = threadMinDiam / 2;
            double len1 = length * 0.1;
            double len2 = length * 0.85;
            double len3 = length * 0.05;

            ksEntity ksSketchE1 = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef1 = ksSketchE1.GetDefinition();

            ksSketchDef1.SetPlane(basePlaneZOY);
            ksSketchE1.Create();
            ksDocument2D Sketch2D1 = (ksDocument2D)ksSketchDef1.BeginEdit();

            Sketch2D1.ksCircle(0, 0, rad1, 1);

            ksSketchDef1.EndEdit();

            ksEntity bossExtr1 = part.NewEntity((short)Obj3dType.o3d_baseExtrusion);
            ksBaseExtrusionDefinition extrDef1 = bossExtr1.GetDefinition();
            ksExtrusionParam extrProp1 = (ksExtrusionParam)extrDef1.ExtrusionParam();

            if (extrProp1 != null)
            {
                extrDef1.SetSketch(ksSketchE1);

                extrProp1.direction = (short)Direction_Type.dtReverse;
                extrProp1.typeReverse = (short)End_Type.etBlind;
                extrProp1.depthReverse = len1;
                bossExtr1.Create();
            }


            ksEntity plane2 = OffsetPlaneCreation(-len1, basePlaneZOY);
            ksEntity ksSketchE2 = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef2 = ksSketchE2.GetDefinition();

            ksSketchDef2.SetPlane(plane2);
            ksSketchE2.Create();
            ksDocument2D Sketch2D2 = (ksDocument2D)ksSketchDef2.BeginEdit();

            Sketch2D2.ksCircle(0, 0, rad2, 1);

            ksSketchDef2.EndEdit();

            ksEntity bossExtr2 = part.NewEntity((short)Obj3dType.o3d_baseExtrusion);
            ksBaseExtrusionDefinition extrDef2 = bossExtr2.GetDefinition();
            ksExtrusionParam extrProp2 = (ksExtrusionParam)extrDef2.ExtrusionParam();

            if (extrProp2 != null)
            {
                extrDef2.SetSketch(ksSketchE2);

                extrProp2.direction = (short)Direction_Type.dtReverse;
                extrProp2.typeReverse = (short)End_Type.etBlind;
                extrProp2.depthReverse = len2;
                bossExtr2.Create();
            }

            ksEntity plane3 = OffsetPlaneCreation(- len1 - len2, basePlaneZOY);
            ksEntity ksSketchE3 = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef3 = ksSketchE3.GetDefinition();

            ksSketchDef3.SetPlane(plane3);
            ksSketchE3.Create();
            ksDocument2D Sketch2D3 = (ksDocument2D)ksSketchDef3.BeginEdit();

            Sketch2D3.ksCircle(0, 0, rad1, 1);

            ksSketchDef3.EndEdit();

            ksEntity bossExtr3 = part.NewEntity((short)Obj3dType.o3d_baseExtrusion);
            ksBaseExtrusionDefinition extrDef3 = bossExtr3.GetDefinition();
            ksExtrusionParam extrProp3 = (ksExtrusionParam)extrDef3.ExtrusionParam();

            if (extrProp3 != null)
            {
                extrDef3.SetSketch(ksSketchE3);

                extrProp3.direction = (short)Direction_Type.dtReverse;
                extrProp3.typeReverse = (short)End_Type.etBlind;
                extrProp3.depthReverse = len3;
                extrProp3.draftOutwardReverse = true;
                extrProp3.draftValueReverse = 45;
                bossExtr3.Create();
            }

        }
        
        ///////////////////////////Создание сквозного отверстия 1 (тип 2 исполнение 1)/////////////////////////////
        private void HoleType2Creation1(double thredMaxDiam, double threadLength1, double threadLength2, double fullLength)
        {
            ksEntity basePlaneZOY = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeYOZ);

            double rad1 = thredMaxDiam / 2;
            double rad2 = rad1 * 3 / 5;
            double rad3 = rad2;
            double rad4 = rad2;
            double rad5 = rad4;
            double rad6 = rad1;
            double len1 = threadLength1;
            double len2 = len1 / 20;
            double len3 = 8;
            double len4 = fullLength + threadLength2 / 2;
            double len5 = 8;
            double len6 = len4 - len1 - len2 - len3 - threadLength2 - len5;

            ksEntity plane1 = OffsetPlaneCreation(0, basePlaneZOY);
            ksEntity ksSketchE1 = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef1 = ksSketchE1.GetDefinition();

            ksSketchDef1.SetPlane(plane1);
            ksSketchE1.Create();
            ksDocument2D Sketch2D1 = (ksDocument2D)ksSketchDef1.BeginEdit();

            Sketch2D1.ksCircle(0, 0, rad1, 1);
            
/*            ksRectangleParam rect1 = (ksRectangleParam)kompas.GetParamStruct((short)StructType2DEnum.ko_RectangleParam);
            if (rect1 != null)
            {
                // Параметры прямоугольника
                rect1.ang = 0;
                rect1.x = -thick;
                rect1.y = rad;
                rect1.width = thick;
                rect1.height = sDiam / 2 - rad;
                rect1.style = 1;
                Sketch2D1.ksRectangle(rect1);
            }*/

            ksSketchDef1.EndEdit();

            ksEntity cutExtr1 = part.NewEntity((short)Obj3dType.o3d_cutExtrusion);
            ksCutExtrusionDefinition extrDef1 = cutExtr1.GetDefinition();
            ksExtrusionParam extrProp1 = (ksExtrusionParam)extrDef1.ExtrusionParam();

            if (extrProp1 != null)
            {
                extrDef1.SetSketch(ksSketchE1);

                extrProp1.direction = (short)Direction_Type.dtReverse;
                extrProp1.typeReverse = (short)End_Type.etBlind;
                extrProp1.depthReverse = len1;
                cutExtr1.Create();
            }

            ksEntity plane2 = OffsetPlaneCreation(len1, basePlaneZOY);
            ksEntity ksSketchE2 = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef2 = ksSketchE2.GetDefinition();

            ksSketchDef2.SetPlane(plane2);
            ksSketchE2.Create();
            ksDocument2D Sketch2D2 = (ksDocument2D)ksSketchDef2.BeginEdit();

            Sketch2D2.ksCircle(0, 0, rad2, 1);

            ksSketchDef2.EndEdit();

            ksEntity cutExtr2 = part.NewEntity((short)Obj3dType.o3d_cutExtrusion);
            ksCutExtrusionDefinition extrDef2 = cutExtr2.GetDefinition();
            ksExtrusionParam extrProp2 = (ksExtrusionParam)extrDef2.ExtrusionParam();

            if (extrProp2 != null)
            {
                extrDef2.SetSketch(ksSketchE2);

                extrProp2.direction = (short)Direction_Type.dtReverse;
                extrProp2.typeReverse = (short)End_Type.etBlind;
                extrProp2.depthReverse = len2;
                cutExtr2.Create();
            }

            ksEntity plane3 = OffsetPlaneCreation(len1 + len2, basePlaneZOY);
            ksEntity ksSketchE3 = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef3 = ksSketchE3.GetDefinition();

            ksSketchDef3.SetPlane(plane3);
            ksSketchE3.Create();
            ksDocument2D Sketch2D3 = (ksDocument2D)ksSketchDef3.BeginEdit();

            Sketch2D3.ksCircle(0, 0, rad3, 1);

            ksSketchDef3.EndEdit();

            ksEntity cutExtr3 = part.NewEntity((short)Obj3dType.o3d_cutExtrusion);
            ksCutExtrusionDefinition extrDef3 = cutExtr3.GetDefinition();
            ksExtrusionParam extrProp3 = (ksExtrusionParam)extrDef3.ExtrusionParam();

            if (extrProp3 != null)
            {
                extrDef3.SetSketch(ksSketchE3);

                extrProp3.direction = (short)Direction_Type.dtReverse;
                extrProp3.typeReverse = (short)End_Type.etBlind;
                extrProp3.depthReverse = len3;
                extrProp3.draftOutwardNormal = true;
                extrProp3.draftValueReverse = 45;
                cutExtr3.Create();
            }

            ksEntity plane4 = OffsetPlaneCreation(len4, basePlaneZOY);
            ksEntity ksSketchE4 = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef4 = ksSketchE4.GetDefinition();

            ksSketchDef4.SetPlane(plane4);
            ksSketchE4.Create();
            ksDocument2D Sketch2D4 = (ksDocument2D)ksSketchDef4.BeginEdit();

            Sketch2D4.ksCircle(0, 0, rad4, 1);

            ksSketchDef4.EndEdit();

            ksEntity cutExtr4 = part.NewEntity((short)Obj3dType.o3d_cutExtrusion);
            ksCutExtrusionDefinition extrDef4 = cutExtr4.GetDefinition();
            ksExtrusionParam extrProp4 = (ksExtrusionParam)extrDef4.ExtrusionParam();

            if (extrProp4 != null)
            {
                extrDef4.SetSketch(ksSketchE4);

                extrProp4.direction = (short)Direction_Type.dtNormal;
                extrProp4.typeNormal = (short)End_Type.etBlind;
                extrProp4.depthNormal = threadLength2;
                cutExtr4.Create();
            }

            ksEntity plane5 = OffsetPlaneCreation(len4 - threadLength2, basePlaneZOY);
            ksEntity ksSketchE5 = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef5 = ksSketchE5.GetDefinition();

            ksSketchDef5.SetPlane(plane5);
            ksSketchE5.Create();
            ksDocument2D Sketch2D5 = (ksDocument2D)ksSketchDef5.BeginEdit();

            Sketch2D5.ksCircle(0, 0, rad5, 1);

            ksSketchDef5.EndEdit();

            ksEntity cutExtr5 = part.NewEntity((short)Obj3dType.o3d_cutExtrusion);
            ksCutExtrusionDefinition extrDef5 = cutExtr5.GetDefinition();
            ksExtrusionParam extrProp5 = (ksExtrusionParam)extrDef5.ExtrusionParam();

            if (extrProp5 != null)
            {
                extrDef5.SetSketch(ksSketchE5);

                extrProp5.direction = (short)Direction_Type.dtNormal;
                extrProp5.typeNormal = (short)End_Type.etBlind;
                extrProp5.depthNormal = len5;
                extrProp5.draftOutwardReverse = true;
                extrProp5.draftValueNormal = 45;
                cutExtr5.Create();
            }

            ksEntity plane6 = OffsetPlaneCreation(len4 - threadLength2 - len5, basePlaneZOY);
            ksEntity ksSketchE6 = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef6 = ksSketchE6.GetDefinition();

            ksSketchDef6.SetPlane(plane6);
            ksSketchE6.Create();
            ksDocument2D Sketch2D6 = (ksDocument2D)ksSketchDef6.BeginEdit();

            Sketch2D6.ksCircle(0, 0, rad6, 1);

            ksSketchDef6.EndEdit();

            ksEntity cutExtr6 = part.NewEntity((short)Obj3dType.o3d_cutExtrusion);
            ksCutExtrusionDefinition extrDef6 = cutExtr6.GetDefinition();
            ksExtrusionParam extrProp6 = (ksExtrusionParam)extrDef6.ExtrusionParam();

            if (extrProp6 != null)
            {
                extrDef6.SetSketch(ksSketchE6);

                extrProp6.direction = (short)Direction_Type.dtNormal;
                extrProp6.typeNormal = (short)End_Type.etBlind;
                extrProp6.depthNormal = len6;
                cutExtr6.Create();
            }
        }

        ///////////////////////////Создание сквозного отверстия 2 (тип 2 исполнение 2)/////////////////////////////
        private void HoleType2Creation2(double threadMaxDiam, double threadLength, double fullLength)
        {
            ksEntity basePlaneZOY = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeYOZ);
            
            double jointLength = threadLength * 4 / 3;
            double rad1 = threadMaxDiam / 3;
            double rad2 = threadMaxDiam / 2;
            double rad3 = rad2 * 7 / 8;
            double rad4 = rad1 * 1.6;
            double len1 = jointLength * 0.05;
            double len2 = jointLength * 0.9;
            double len3 = jointLength * 0.1;
            double len4 = threadLength;
            double len5 = threadLength / 2;
            double len6 = len5 * 1 / 6;
            double len7 = fullLength - len6 - len5 - len4 - jointLength * 0.05;

            ksEntity plane1 = OffsetPlaneCreation(-jointLength + len1, basePlaneZOY);
            ksEntity ksSketchE1 = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef1 = ksSketchE1.GetDefinition();

            ksSketchDef1.SetPlane(plane1);
            ksSketchE1.Create();
            ksDocument2D Sketch2D1 = (ksDocument2D)ksSketchDef1.BeginEdit();

            Sketch2D1.ksCircle(0, 0, rad1, 1);

            ksSketchDef1.EndEdit();

            ksEntity cutExtr1 = part.NewEntity((short)Obj3dType.o3d_cutExtrusion);
            ksCutExtrusionDefinition extrDef1 = cutExtr1.GetDefinition();
            ksExtrusionParam extrProp1 = (ksExtrusionParam)extrDef1.ExtrusionParam();

            if (extrProp1 != null)
            {
                extrDef1.SetSketch(ksSketchE1);

                extrProp1.direction = (short)Direction_Type.dtNormal;
                extrProp1.typeNormal = (short)End_Type.etBlind;
                extrProp1.depthNormal = len1;
                extrProp1.draftOutwardReverse = true;
                extrProp1.draftValueNormal = 45;
                cutExtr1.Create();
            }

            ksEntity plane2 = OffsetPlaneCreation(-jointLength + len1, basePlaneZOY);
            ksEntity ksSketchE2 = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef2 = ksSketchE2.GetDefinition();

            ksSketchDef2.SetPlane(plane2);
            ksSketchE2.Create();
            ksDocument2D Sketch2D2 = (ksDocument2D)ksSketchDef2.BeginEdit();

            Sketch2D2.ksCircle(0, 0, rad1, 1);

            ksSketchDef2.EndEdit();

            ksEntity cutExtr2 = part.NewEntity((short)Obj3dType.o3d_cutExtrusion);
            ksCutExtrusionDefinition extrDef2 = cutExtr2.GetDefinition();
            ksExtrusionParam extrProp2 = (ksExtrusionParam)extrDef2.ExtrusionParam();

            if (extrProp2 != null)
            {
                extrDef2.SetSketch(ksSketchE2);

                extrProp2.direction = (short)Direction_Type.dtReverse;
                extrProp2.typeReverse = (short)End_Type.etBlind;
                extrProp2.depthReverse = len2;
                cutExtr2.Create();
            }

            ksEntity plane3 = OffsetPlaneCreation(-jointLength + len1 + len2, basePlaneZOY);
            ksEntity ksSketchE3 = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef3 = ksSketchE3.GetDefinition();

            ksSketchDef3.SetPlane(plane3);
            ksSketchE3.Create();
            ksDocument2D Sketch2D3 = (ksDocument2D)ksSketchDef3.BeginEdit();

            Sketch2D3.ksCircle(0, 0, rad1, 1);

            ksSketchDef3.EndEdit();

            ksEntity cutExtr3 = part.NewEntity((short)Obj3dType.o3d_cutExtrusion);
            ksCutExtrusionDefinition extrDef3 = cutExtr3.GetDefinition();
            ksExtrusionParam extrProp3 = (ksExtrusionParam)extrDef3.ExtrusionParam();

            if (extrProp3 != null)
            {
                extrDef3.SetSketch(ksSketchE3);

                extrProp3.direction = (short)Direction_Type.dtReverse;
                extrProp3.typeReverse = (short)End_Type.etBlind;
                extrProp3.depthReverse = len3;
                extrProp3.draftOutwardNormal = true;
                extrProp3.draftValueReverse = 30;
                cutExtr3.Create();
            }

            ksEntity plane4 = OffsetPlaneCreation(fullLength, basePlaneZOY);
            ksEntity ksSketchE4 = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef4 = ksSketchE4.GetDefinition();

            ksSketchDef4.SetPlane(plane4);
            ksSketchE4.Create();
            ksDocument2D Sketch2D4 = (ksDocument2D)ksSketchDef4.BeginEdit();

            Sketch2D4.ksCircle(0, 0, rad2, 1);

            ksSketchDef4.EndEdit();

            ksEntity cutExtr4 = part.NewEntity((short)Obj3dType.o3d_cutExtrusion);
            ksCutExtrusionDefinition extrDef4 = cutExtr4.GetDefinition();
            ksExtrusionParam extrProp4 = (ksExtrusionParam)extrDef4.ExtrusionParam();

            if (extrProp4 != null)
            {
                extrDef4.SetSketch(ksSketchE4);

                extrProp4.direction = (short)Direction_Type.dtNormal;
                extrProp4.typeNormal = (short)End_Type.etBlind;
                extrProp4.depthNormal = len4;
                cutExtr4.Create();
            }

            ksEntity plane5 = OffsetPlaneCreation(fullLength - len4, basePlaneZOY);
            ksEntity ksSketchE5 = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef5 = ksSketchE5.GetDefinition();

            ksSketchDef5.SetPlane(plane5);
            ksSketchE5.Create();
            ksDocument2D Sketch2D5 = (ksDocument2D)ksSketchDef5.BeginEdit();

            Sketch2D5.ksCircle(0, 0, rad3, 1);

            ksSketchDef5.EndEdit();

            ksEntity cutExtr5 = part.NewEntity((short)Obj3dType.o3d_cutExtrusion);
            ksCutExtrusionDefinition extrDef5 = cutExtr5.GetDefinition();
            ksExtrusionParam extrProp5 = (ksExtrusionParam)extrDef5.ExtrusionParam();

            if (extrProp5 != null)
            {
                extrDef5.SetSketch(ksSketchE5);

                extrProp5.direction = (short)Direction_Type.dtNormal;
                extrProp5.typeNormal = (short)End_Type.etBlind;
                extrProp5.depthNormal = len5;
                cutExtr5.Create();
            }

            ksEntity plane6 = OffsetPlaneCreation(fullLength - len4 - len5, basePlaneZOY);
            ksEntity ksSketchE6 = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef6 = ksSketchE6.GetDefinition();

            ksSketchDef6.SetPlane(plane6);
            ksSketchE6.Create();
            ksDocument2D Sketch2D6 = (ksDocument2D)ksSketchDef6.BeginEdit();

            Sketch2D6.ksCircle(0, 0, rad3, 1);

            ksSketchDef6.EndEdit();

            ksEntity cutExtr6 = part.NewEntity((short)Obj3dType.o3d_cutExtrusion);
            ksCutExtrusionDefinition extrDef6 = cutExtr6.GetDefinition();
            ksExtrusionParam extrProp6 = (ksExtrusionParam)extrDef6.ExtrusionParam();

            if (extrProp6 != null)
            {
                extrDef6.SetSketch(ksSketchE6);

                extrProp6.direction = (short)Direction_Type.dtNormal;
                extrProp6.typeNormal = (short)End_Type.etBlind;
                extrProp6.depthNormal = len6;
                extrProp6.draftOutwardReverse = true;
                extrProp6.draftValueNormal = 30;
                cutExtr6.Create();
            }

            ksEntity plane7 = OffsetPlaneCreation(fullLength - len4 - len5 - len6, basePlaneZOY);
            ksEntity ksSketchE7 = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef7 = ksSketchE7.GetDefinition();

            ksSketchDef7.SetPlane(plane7);
            ksSketchE7.Create();
            ksDocument2D Sketch2D7 = (ksDocument2D)ksSketchDef7.BeginEdit();

            Sketch2D7.ksCircle(0, 0, rad4, 1);

            ksSketchDef7.EndEdit();

            ksEntity cutExtr7 = part.NewEntity((short)Obj3dType.o3d_cutExtrusion);
            ksCutExtrusionDefinition extrDef7 = cutExtr7.GetDefinition();
            ksExtrusionParam extrProp7 = (ksExtrusionParam)extrDef7.ExtrusionParam();

            if (extrProp7 != null)
            {
                extrDef7.SetSketch(ksSketchE7);

                extrProp7.direction = (short)Direction_Type.dtNormal;
                extrProp7.typeNormal = (short)End_Type.etBlind;
                extrProp7.depthNormal = len7;
                cutExtr7.Create();
            }
        }

        ///////////////////////////Создание кастомных плоскостей/////////////////////////////
        private ksEntity OffsetPlaneCreation(double distance, ksEntity plane)
        {
            ksEntity basePlaneOffset = (ksEntity)part.NewEntity((short)Obj3dType.o3d_planeOffset);
            PlaneOffsetDefinition offsetPlaneDef = basePlaneOffset.GetDefinition();
            offsetPlaneDef.direction = true;
            offsetPlaneDef.offset = distance;
            offsetPlaneDef.SetPlane(plane);
            basePlaneOffset.hidden = true;
            basePlaneOffset.Create();

            return basePlaneOffset;
        }

        ///////////////////////////Создание винта/////////////////////////////
        private void SpyralCreation(double rad, double spyralStep, double start, double spyralLength, double thick, double sDiam)
        {
            ksEntity basePlaneZOY = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeYOZ);
            ksEntity basePlaneXOZ = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOZ);
            ksEntity startPlane = OffsetPlaneCreation(start, basePlaneZOY);

            //траектория
            ksEntity ksSketchE1 = part.NewEntity((short)Obj3dType.o3d_cylindricSpiral);

            CylindricSpiralDefinition ksSketchDef1 = ksSketchE1.GetDefinition();

            ksSketchDef1.SetPlane(startPlane);

            ksSketchE1.hidden = true;
            ksSketchDef1.diam = rad * 2;
            ksSketchDef1.buildMode = 0;
            ksSketchDef1.step = spyralStep;
            ksSketchDef1.turn = spyralLength / spyralStep;
            ksSketchDef1.buildDir = true;
            ksSketchDef1.turnDir = true;

            ksSketchE1.Create();

            //выдавливаемый профиль
            ksEntity ksSketchE2 = part.NewEntity((short)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef2 = ksSketchE2.GetDefinition();

            ksSketchDef2.SetPlane(basePlaneXOZ);
            ksSketchE2.hidden = true;
            ksSketchE2.Create();
            ksDocument2D Sketch2D2 = (ksDocument2D)ksSketchDef2.BeginEdit();

            ksRectangleParam rect = (ksRectangleParam)kompas.GetParamStruct((short)StructType2DEnum.ko_RectangleParam);
            if (rect != null)
            {
                // Параметры прямоугольника
                rect.ang = 0;
                rect.x = -thick - start;
                rect.y = rad;
                rect.width = thick;
                rect.height = sDiam / 2 - rad;
                rect.style = 1;
                Sketch2D2.ksRectangle(rect);
            }

            ksSketchDef2.EndEdit();

            //выдавливание профиля по траектории
            ksEntity trajectoryExtr = part.NewEntity((short)Obj3dType.o3d_baseEvolution);
            ksBaseEvolutionDefinition extrDef = trajectoryExtr.GetDefinition();

            extrDef.PathPartArray().add(ksSketchE1);
            extrDef.SetSketch(ksSketchE2);
            trajectoryExtr.Create();
        }

        ///////////////////////////Создание конуса экструзионного шнека/////////////////////////////
        private void ConeCreation(double rad)
        {
            ksEntity basePlaneZOY = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeYOZ);

            double length = rad * 100;

            ksEntity ksSketchE = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef = ksSketchE.GetDefinition();

            ksSketchDef.SetPlane(basePlaneZOY);
            ksSketchE.Create();
            ksDocument2D Sketch2D = (ksDocument2D)ksSketchDef.BeginEdit();

            Sketch2D.ksCircle(0, 0, rad, 1);

            ksSketchDef.EndEdit();

            ksEntity baseExtr = part.NewEntity((short)Obj3dType.o3d_baseExtrusion);
            ksBaseExtrusionDefinition extrDef = baseExtr.GetDefinition();
            ksExtrusionParam extrProp = (ksExtrusionParam)extrDef.ExtrusionParam();

            if (extrProp != null)
            {
                extrDef.SetSketch(ksSketchE);

                extrProp.direction = (short)Direction_Type.dtReverse;
                extrProp.typeReverse = (short)End_Type.etBlind;
                extrProp.depthReverse = length;
                extrProp.draftOutwardReverse = true;
                extrProp.draftValueReverse = 30;
                baseExtr.Create();
            }
        }

        private void ParamConv()
        {
            mistakeCheck = true;

            if (mistakeCheck == true)
            {
                foreach (TextBox textBox in FindVisualChildren<TextBox>(System.Windows.Application.Current.MainWindow))
                {
                    textBox.BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(171, 173, 179));
                }
            }
            if (string.IsNullOrEmpty(inputTubeLength.Text))
            {
                mistakeCheck = false;
                inputTubeLength.BorderBrush = Brushes.Red;
                MessageBox.Show("Обнаружено пустое поле ввода!");
            }
            else if (string.IsNullOrEmpty(inputShnekDiam.Text))
            {
                mistakeCheck = false;
                inputShnekDiam.BorderBrush = Brushes.Red;
                MessageBox.Show("Обнаружено пустое поле ввода!");
            }
            else if (string.IsNullOrEmpty(inputHoleDiam.Text))
            {
                mistakeCheck = false;
                inputHoleDiam.BorderBrush = Brushes.Red;
                MessageBox.Show("Обнаружено пустое поле ввода!");
            }
            else if (string.IsNullOrEmpty(inputHexSize.Text))
            {
                mistakeCheck = false;
                inputHexSize.BorderBrush = Brushes.Red;
                MessageBox.Show("Обнаружено пустое поле ввода!");
            }
            else if (string.IsNullOrEmpty(inputHoleDistance.Text))
            {
                mistakeCheck = false;
                inputHoleDistance.BorderBrush = Brushes.Red;
                MessageBox.Show("Обнаружено пустое поле ввода!");
            }
            else if (string.IsNullOrEmpty(inputHex2Size.Text))
            {
                mistakeCheck = false;
                inputHex2Size.BorderBrush = Brushes.Red;
                MessageBox.Show("Обнаружено пустое поле ввода!");
            }
            else if (string.IsNullOrEmpty(inputStep.Text))
            {
                mistakeCheck = false;
                inputStep.BorderBrush = Brushes.Red;
                MessageBox.Show("Обнаружено пустое поле ввода!");
            }
            else if (string.IsNullOrEmpty(inputShnekThick.Text))
            {
                mistakeCheck = false;
                inputShnekThick.BorderBrush = Brushes.Red;
                MessageBox.Show("Обнаружено пустое поле ввода!");
            }
            else if (string.IsNullOrEmpty(inputType2ShnekDiam.Text))
            {
                mistakeCheck = false;
                inputType2ShnekDiam.BorderBrush = Brushes.Red;
                MessageBox.Show("Обнаружено пустое поле ввода!");
            }
            else if (string.IsNullOrEmpty(inputType2T.Text))
            {
                mistakeCheck = false;
                inputType2T.BorderBrush = Brushes.Red;
                MessageBox.Show("Обнаружено пустое поле ввода!");
            }
            else if (string.IsNullOrEmpty(inputType2T1.Text))
            {
                mistakeCheck = false;
                inputType2T1.BorderBrush = Brushes.Red;
                MessageBox.Show("Обнаружено пустое поле ввода!");
            }
            else if (string.IsNullOrEmpty(inputType2T2.Text))
            {
                mistakeCheck = false;
                inputType2T2.BorderBrush = Brushes.Red;
                MessageBox.Show("Обнаружено пустое поле ввода!");
            }
            else if (string.IsNullOrEmpty(inputThreadDiam.Text))
            {
                mistakeCheck = false;
                inputThreadDiam.BorderBrush = Brushes.Red;
                MessageBox.Show("Обнаружено пустое поле ввода!");
            }
            else if (string.IsNullOrEmpty(inputThreadStep.Text))
            {
                mistakeCheck = false;
                inputThreadStep.BorderBrush = Brushes.Red;
                MessageBox.Show("Обнаружено пустое поле ввода!");
            }
            else if (string.IsNullOrEmpty(inputExtrShnekDiam.Text))
            {
                mistakeCheck = false;
                inputExtrShnekDiam.BorderBrush = Brushes.Red;
                MessageBox.Show("Обнаружено пустое поле ввода!");
            }
            else if (string.IsNullOrEmpty(inputExtrShnekCoffLength.Text))
            {
                mistakeCheck = false;
                inputExtrShnekCoffLength.BorderBrush = Brushes.Red;
                MessageBox.Show("Обнаружено пустое поле ввода!");
            }
            else
            {
                //mistakeCheck = true;
                tubeLength = Convert.ToDouble(inputTubeLength.Text);
                shnekDiam = Convert.ToDouble(inputShnekDiam.Text);
                holeDiam = Convert.ToDouble(inputHoleDiam.Text);
                hexSize = Convert.ToDouble(inputHexSize.Text);
                holeDistance = Convert.ToDouble(inputHoleDistance.Text);
                hex2Size = Convert.ToDouble(inputHex2Size.Text);
                step = Convert.ToDouble(inputStep.Text);
                shnekThick = Convert.ToDouble(inputShnekThick.Text);

                type2ShnekDiam = Convert.ToDouble(inputType2ShnekDiam.Text);
                type2T = Convert.ToDouble(inputType2T.Text);
                type2T1 = Convert.ToDouble(inputType2T1.Text);
                type2T2 = Convert.ToDouble(inputType2T2.Text);
                threadDiam = Convert.ToDouble(inputThreadDiam.Text);
                threadStep = Convert.ToDouble(inputThreadStep.Text);

                extrDiam = Convert.ToDouble(inputExtrShnekDiam.Text);
                extrCoffLength = Convert.ToDouble(inputExtrShnekCoffLength.Text);

                if (GhostType.SelectedIndex == 0 || GhostType.SelectedIndex == 1)
                {
                    if (tubeLength < 1000 || tubeLength > 2500)
                    {
                        inputTubeLength.BorderBrush = Brushes.Red;
                        MessageBox.Show("Длина шнека должна находиться в диапазоне от 1000 до 2500 мм!");
                        mistakeCheck = false;
                    }
                }

                if (GhostType.SelectedIndex == 0 || GhostType.SelectedIndex == 2) mistakeCheck = true;
                else
                {
                    //буровые шнеки
                    if (GhostType.SelectedIndex == 1)
                    {
                        //первый тип
                        if (ShnekType.SelectedIndex == 0)
                        {
                            if (holeDiam == 0)
                            {
                                inputHoleDiam.BorderBrush = Brushes.Red;
                                MessageBox.Show("Введён неверный диаметр отверстия!");
                                mistakeCheck = false;
                            }
                            if (holeDistance == 0)
                            {
                                inputHoleDistance.BorderBrush = Brushes.Red;
                                MessageBox.Show("Введено неверное расстояние отверстия!");
                                mistakeCheck = false;
                            }
                            if (shnekDiam == 0)
                            {
                                inputShnekDiam.BorderBrush = Brushes.Red;
                                MessageBox.Show("Введён неверный внешний диаметр шнека!");
                                mistakeCheck = false;
                            }
                            if (step == 0)
                            {
                                inputShnekDiam.BorderBrush = Brushes.Red;
                                MessageBox.Show("Введён неверный внешний диаметр шнека!");
                                mistakeCheck = false;
                            }
                            if (shnekThick == 0)
                            {
                                inputShnekDiam.BorderBrush = Brushes.Red;
                                MessageBox.Show("Введён неверный внешний диаметр шнека!");
                                mistakeCheck = false;
                            }
                            if (hexSize * 1.5 >= shnekDiam)
                            {
                                inputShnekDiam.BorderBrush = Brushes.Red;
                                MessageBox.Show("Внешний диаметр шнека не может быть меньше или равен внутреннему!");
                                mistakeCheck = false;
                            }
                        }
                        //второй тип
                        else if (ShnekType.SelectedIndex == 1)
                        {
                            if (type2T1 < 100)
                            {
                                inputType2T1.BorderBrush = Brushes.Red;
                                MessageBox.Show("Параметр t1 не может быть меньше 100 мм!");
                                mistakeCheck = false;
                            }
                            if (type2T1 >= tubeLength*0.3)
                            {
                                inputType2T1.BorderBrush = Brushes.Red;
                                MessageBox.Show("Параметр t1 не может быть превышать 30% от длины трубы!");
                                mistakeCheck = false;
                            }
                            if (type2T2 < 200)
                            {
                                inputType2T2.BorderBrush = Brushes.Red;
                                MessageBox.Show("Параметр t2 не может быть меньше 200 мм!");
                                mistakeCheck = false;
                            }
                            if (type2T2 >= tubeLength*0.6)
                            {
                                inputType2T2.BorderBrush = Brushes.Red;
                                MessageBox.Show("Параметр t2 не может быть превышать 60% от длины трубы!");
                                mistakeCheck = false;
                            }
                            if (threadDiam >= type2ShnekDiam / 2)
                            {
                                inputThreadDiam.BorderBrush = Brushes.Red;
                                MessageBox.Show("Диаметр отверстия не может быть больше или равен диаметру трубы!");
                                mistakeCheck = false;
                            }
                            if (type2T < 30 || type2T > 100)
                            {
                                inputType2T.BorderBrush = Brushes.Red;
                                MessageBox.Show("Параметр t должен находиться в диапазоне от 30 до 100 мм!");
                                mistakeCheck = false;
                            }
                            if (threadStep < 4 || threadStep > 20)
                            {
                                inputThreadStep.BorderBrush = Brushes.Red;
                                MessageBox.Show("Шаг резьбы должен находиться в диапазоне от 4 до 20 мм!");
                                mistakeCheck = false;
                            }
                        }
                    }
                    //экструзионные шнеки
                    if (GhostType.SelectedIndex == 3)
                    {
                        if (extrDiam == 0)
                        {
                            inputExtrShnekDiam.BorderBrush = Brushes.Red;
                            MessageBox.Show("Введён неверный диаметр экструзионного шнека!");
                            mistakeCheck = false;
                        }
                        if (ShnekType.SelectedIndex == 0)
                        {
                            if (extrCoffLength < 20 || extrCoffLength > 30)
                            {
                                inputExtrShnekCoffLength.BorderBrush = Brushes.Red;
                                MessageBox.Show("Отношение длины экструзионного шнека к его диаметру должно находиться в диапазоне от 20 до 30!");
                                mistakeCheck = false;
                            }
                        }
                        if (ShnekType.SelectedIndex == 1)
                        {
                            if (extrCoffLength < 6 || extrCoffLength > 12)
                            {
                                inputExtrShnekCoffLength.BorderBrush = Brushes.Red;
                                MessageBox.Show("Отношение длины экструзионного шнека к его диаметру должно находиться в диапазоне от 6 до 12!");
                                mistakeCheck = false;
                            }
                        }
                    }
                    if (ShnekType.SelectedIndex == 0)
                    {
                        if (ShnekStyle.SelectedIndex == 0)
                        {
                            if (hexSize == 0)
                            {
                                inputHexSize.BorderBrush = Brushes.Red;
                                MessageBox.Show("Введён неверный размер шестигранника!");
                                mistakeCheck = false;
                            }
                            if (holeDiam * 2 >= hexSize)
                            {
                                inputHoleDiam.BorderBrush = Brushes.Red;
                                inputHexSize.BorderBrush = Brushes.Red;
                                MessageBox.Show("Диаметр отверстия не может быть больше боковой грани шестигранника!");
                                mistakeCheck = false;
                            }
                        }
                        else
                        {
                            if (hex2Size == 0)
                            {
                                inputShnekDiam.BorderBrush = Brushes.Red;
                                MessageBox.Show("Введён неверный внешний диаметр шнека!");
                                mistakeCheck = false;
                            }
                            if (holeDiam * 55 / 24 >= hex2Size)
                            {
                                inputHoleDiam.BorderBrush = Brushes.Red;
                                inputHex2Size.BorderBrush = Brushes.Red;
                                MessageBox.Show("Диаметр отверстия не может быть больше боковой грани присоединительного элемента!");
                                mistakeCheck = false;
                            }
                        }
                    }
                    if (holeDiam > holeDistance / 2)
                    {
                        inputHoleDiam.BorderBrush = Brushes.Red;
                        inputHoleDistance.BorderBrush = Brushes.Red;
                        MessageBox.Show("Диаметр отверстия не может превышать длину присоединительного элемента!");
                        mistakeCheck = false;
                    }
                    if (shnekThick >= step)
                    {
                        inputShnekThick.BorderBrush = Brushes.Red;
                        inputStep.BorderBrush = Brushes.Red;
                        MessageBox.Show("Толщина винта должна быть меньше шага!");
                        mistakeCheck = false;
                    }
                    if (step < 50 || step > 200)
                    {
                        inputStep.BorderBrush = Brushes.Red;
                        MessageBox.Show("Шаг винта должен находиться в диапазоне от 50 до 200 мм!");
                        mistakeCheck = false;
                    }
                }
            }
        }

        private void InputFieldIsActive(bool isActive)
        {
            inputHoleDiam.IsEnabled = isActive;
            inputShnekThick.IsEnabled = isActive;
            inputShnekDiam.IsEnabled = isActive;
            inputHexSize.IsEnabled = isActive;
            inputHoleDistance.IsEnabled = isActive;
            inputStep.IsEnabled = isActive;
            inputHex2Size.IsEnabled = isActive;
            inputThreadDiam.IsEnabled = isActive;
            inputThreadStep.IsEnabled = isActive;
            inputType2ShnekDiam.IsEnabled = isActive;
            inputType2T.IsEnabled = isActive;
            inputType2T1.IsEnabled = isActive;
            inputType2T2.IsEnabled = isActive;
            inputExtrShnekDiam.IsEnabled = isActive;
            inputExtrShnekCoffLength.IsEnabled = isActive;
        }

        private void InputFieldIvVisible(bool isVisible)
        {
            if (isVisible)
            {
                inputSelection1.Visibility = Visibility.Visible;
                inputSelection2.Visibility = Visibility.Collapsed;
            }
            else
            {
                inputSelection1.Visibility = Visibility.Collapsed;
                inputSelection2.Visibility = Visibility.Visible;
            }
        }

        private void GOSTSelection1()
        {
            //ShnekType.IsEnabled = true;
            typeCount = 2;
            InputFieldIsActive(false);
            inputTubeLength.IsEnabled = true;
            ShnekStyle.IsEnabled = false;
            ShnekType.IsEnabled = true;
            DefaultShnekChoose.IsEnabled = true;
            InputFieldIvVisible(true);
            for (int i = 0; i < typeCount; i++) ShnekType.Items.Add($"Тип {i + 1}");
            ShnekType.SelectedIndex = 0;
        }

        private void GOSTSelection2()
        {
            typeCount = 2;
            InputFieldIsActive(true);
            inputTubeLength.IsEnabled = true;
            //ShnekType.IsEnabled = false;
            ShnekStyle.IsEnabled = true;
            ShnekType.IsEnabled = true;
            DefaultShnekChoose.IsEnabled = false;
            InputFieldIvVisible(true);
            for (int i = 0; i < typeCount; i++) ShnekType.Items.Add($"Тип {i + 1}");
            ShnekType.SelectedIndex = 0;
        }

        private void GOSTSelection3()
        {
            typeCount = 2;
            InputFieldIsActive(false);
            ShnekStyle.IsEnabled = false;
            ShnekStyle.Items.Clear();
            ShnekType.IsEnabled = true;
            DefaultShnekChoose.IsEnabled = true;
            DefaultShnekItems3();
            InputFieldIvVisible(false);
            for (int i = 0; i < typeCount; i++)
            {
                if (i == 0 ) ShnekType.Items.Add("Для термопластов");
                if (i == 1 ) ShnekType.Items.Add("Для резиновых смесей");
            }
            ShnekType.SelectedIndex = 0;
            if (ImgSketch != null) ImgSketch.Source = (ImageSource)new ImageSourceConverter().ConvertFrom(new Uri(@"D:\Users\Garnik\Desktop\учёба\Диплом\GenShnekApp\GenShnekApp\ShnekSketch3.png"));
            if (ImgTable != null) ImgTable.Source = (ImageSource)new ImageSourceConverter().ConvertFrom(new Uri(@"D:\Users\Garnik\Desktop\учёба\Диплом\GenShnekApp\GenShnekApp\ShnekTable3.png"));
        }
        private void GOSTSelection4()
        {
            typeCount = 2;
            InputFieldIsActive(true);
            ShnekStyle.IsEnabled = false;
            ShnekStyle.Items.Clear();
            ShnekType.IsEnabled = true;
            DefaultShnekChoose.IsEnabled = false;
            DefaultShnekItems3();
            InputFieldIvVisible(false);
            for (int i = 0; i < typeCount; i++)
            {
                if (i == 0) ShnekType.Items.Add("Для пластических масс");
                if (i == 1) ShnekType.Items.Add("Для резиновых смесей");
            }
            ShnekType.SelectedIndex = 0;
            if (ImgSketch != null) ImgSketch.Source = (ImageSource)new ImageSourceConverter().ConvertFrom(new Uri(@"D:\Users\Garnik\Desktop\учёба\Диплом\GenShnekApp\GenShnekApp\ShnekSketch3.png"));
            if (ImgTable != null) ImgTable.Source = (ImageSource)new ImageSourceConverter().ConvertFrom(new Uri(@"D:\Users\Garnik\Desktop\учёба\Диплом\GenShnekApp\GenShnekApp\ShnekTable3.png"));
        }

        private void DefaultShnekItems1()
        {
            for (int i = 0; i < 6; i++)
            {
                if (i == 0)
                {
                    DefaultShnekChoose.Items.Add("ШБ-135");
                }
                if (i == 1)
                {
                    DefaultShnekChoose.Items.Add("ШБ-150");
                }
                if (i == 2)
                {
                    DefaultShnekChoose.Items.Add("ШБ-180");
                }
                if (i == 3)
                {
                    DefaultShnekChoose.Items.Add("ШБ-200");
                }
                if (i == 4)
                {
                    DefaultShnekChoose.Items.Add("ШБ-300");
                }
                if (i == 5)
                {
                    DefaultShnekChoose.Items.Add("ШБ-300У");
                }
            }
            DefaultShnekChoose.SelectedIndex = 0;

            inputSelection11.Visibility = Visibility.Visible;
            inputSelection12.Visibility = Visibility.Collapsed;
        }
        private void DefaultShnekItems2()
        {
            for (int i = 0; i < 3; i++)
            {
                if (i == 0)
                {
                    DefaultShnekChoose.Items.Add("ШС-80");
                }
                if (i == 1)
                {
                    DefaultShnekChoose.Items.Add("ШС-100");
                }
                if (i == 2)
                {
                    DefaultShnekChoose.Items.Add("ШС-200");
                }
            }
            DefaultShnekChoose.SelectedIndex = 0;

            inputSelection11.Visibility = Visibility.Collapsed;
            inputSelection12.Visibility = Visibility.Visible;
        }

        private void DefaultShnekItems3()
        {
            for (int i = 0; i < 13; i++)
            {
                if (i == 0)
                {
                    DefaultShnekChoose.Items.Add("ЧП 20х20");
                }
                if (i == 1)
                {
                    DefaultShnekChoose.Items.Add("ЧП 32х20");
                }
                if (i == 2)
                {
                    DefaultShnekChoose.Items.Add("ЧП 45х20");
                }
                if (i == 3)
                {
                    DefaultShnekChoose.Items.Add("ЧП 45х25");
                }
                if (i == 4)
                {
                    DefaultShnekChoose.Items.Add("ЧП 63х20");
                }
                if (i == 5)
                {
                    DefaultShnekChoose.Items.Add("ЧП 63х25");
                }
                if (i == 6)
                {
                    DefaultShnekChoose.Items.Add("ЧП 63х30");
                }
                if (i == 7)
                {
                    DefaultShnekChoose.Items.Add("ЧП 90х20");
                }
                if (i == 8)
                {
                    DefaultShnekChoose.Items.Add("ЧП 90х25");
                }
                if (i == 9)
                {
                    DefaultShnekChoose.Items.Add("ЧП 90х30");
                }
                if (i == 10)
                {
                    DefaultShnekChoose.Items.Add("ЧП 125х25");
                }
                if (i == 11)
                {
                    DefaultShnekChoose.Items.Add("ЧП 160х20");
                }
                if (i == 12)
                {
                    DefaultShnekChoose.Items.Add("ЧП 200х20");
                }
                DefaultShnekChoose.SelectedIndex = 0;
            }
        }

        //Метод для поиска всех текстбоксов (для изменения цвета)
        public static IEnumerable<T> FindVisualChildren<T>(DependencyObject depObj) where T : DependencyObject
        {
            if (depObj != null)
            {
                for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
                {
                    DependencyObject child = VisualTreeHelper.GetChild(depObj, i);
                    if (child is T)
                    {
                        yield return (T)child;
                    }

                    foreach (T childOfChild in FindVisualChildren<T>(child))
                    {
                        yield return childOfChild;
                    }
                }
            }
        }

        private void CloseButton(object sender, RoutedEventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }

        private void NoteButton(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Расчёты проводятся только для экструзионных шнеков.");
        }

        private void ThreadButton(object sender, RoutedEventArgs e)
        {
            ThreadInfoWindow threadInfoWindow = new ThreadInfoWindow();
            threadInfoWindow.Owner = this;
            threadInfoWindow.ShowDialog();
        }

        //Расчёт шнека на прочность, жёсткость и устойчивость
        ///////////(1)НАЧАЛО
        private void ShnekCalc(double diam, double L)
        {
            //РАСЧЁТ НА ПРОЧНОСТЬ КОНСОЛЬНО ЗАКРЕПЛЁННОГО ШНЕКА
            const double PI = Math.PI; //Число ПИ
            const double gi = 9.81; //ускорение свободного падения

            double A; //постоянная прямого потока
            double B; //постоянная обратного потока
            double G; //поток утечки
            double ZX; //число отрезков разбиения
            double Q; //производительность
            double MKR; //крутящий момент
            double F; //площадь поперечного сечения
            double Sos; //осевое усилие от давления формования
            double K; //параметр
            double J; //момент инерции поперечного сечения

            int size = 401;
            double[] MIZ = new double[size]; //изгибающий момент
            double[] X = new double[size]; //текущая координата по длине шнека, м
            double[] MK = new double[size];
            double[] Fmax1 = new double[size];
            double[] Fmax2 = new double[size];
            double[] Fmax01 = new double[size];
            double[] Fmax02 = new double[size];
            double[] Fmax03 = new double[size];

            double AF = 0;
            double dX = 0.0045;
            double hx1 = 0.0045;
            const double E = 200000000000; //модуль упругости Юнга, Па

            //if (AF == 0)
            //{
            //Входные параметры
            ///////////(2)ГЛАВНЫЕ ПАРАМЕТРЫ
            double SIG = 400000000; //допускаемое напряжение, Па
            double RO = 7850; //плотность материала шнека, кг/м3
            double P = 50000000; //давление развиваемое шнеком, Па
            //TODO просто оставлю тудушку, на случай, если всё-таки будет осевое отверстие. пока оно равно нулю, то есть отверстия нет
            double d1 = 0; //диаметр осевого отверстия шнека, м
            //double d = 0.032; //наружный диаметр шнека, м
            //d = 0.032; //наружный диаметр шнека, м
            diam /= 1000;
            //double L = 0.64; //длина нарезной части шнека, м
            L /= 1000;
            double H = 0.0032; //глубина винтового канала шнека, м
            double FI = 17; //угол наклона винтовой линии шнека, град
            double T = 0.032; //шаг винтовой нарезки шнека, м
            double E1 = 0.0032; //ширина гребня винтового канала шнека, м
            double N = 5; //технологическая мощность, кВт
            double W = 70; //частота вращения шнека, об/мин
            double GAM = RO * gi; //удельный вес материала

            double AL; //альфа, отношение диаметра осевого отверстия шнека к наружному диаметру шнека
            AL = d1 / diam;

            //}

            ///////////(3)ОПРЕДЕЛЕНЕИ ЧИСЛА ОТРЕЗКОВ РАЗБИЕНИЯ ДЛИНЫ НАРЕЗНОЙ ЧАСТИ ШНЕКА
            ZX = L / dX;

            ///////////(4)ОПРЕДЕЛЕНИЕ ПОСТОЯННЫХ ПРЯМОГО, ОБРАТНОГО ПОТОКА, ПОТОКА УТЕЧКИ И ПРОИЗВОДИТЕЛЬНОСТИ
            FI = FI * PI / 180;
            A = PI * diam * H * (T - E) * Pow(Cos(FI), 2) / 2;
            B = Pow(H, 3) * (T - E) * Sin(2 * FI) / (24 * L);
            G = Pow(PI, 2) * Pow(diam, 2) * Pow(diam * L, 3) * Tan(FI) * Sin(FI) / (10 * E1 * L);
            J = ((PI * Pow(diam, 4)) / 64) * (1 - Pow(AL, 4));
            K = Sqrt(P / (E * J));
            Q = A * K * N / (K - B - G); //Вывод
            QOutput.Text = $"Q = {Q} м^3/с";

            ///////////(5)ОПРЕДЕЛЕНИЕ КРУТЯЩЕГО МОМЕНТА, ПЛОЩАДИ ПОПЕРЕЧНОГО СЕЧЕНИЯ ШНЕКА И ОСЕВОГО УСИЛИЯ
            MKR = 9550 * N / W; //Вывод
            MKROutput.Text = $"M_кр = {MKR} Н*м";
            F = PI * Pow(diam, 2) / 4;
            Sos = F * P; //Вывод
            SosOutput.Text = $"S_ос = {Sos} Н";

            ///////////(6)РАСЧЁТ ГИБКОСТИ ШНЕКА
            double F1; // площадь поперечного сечения шнека сечения А-А
            double J1; // момент инерции поперечного сечения А-А
            double I; // радиус инерции сечения
            double MU = 2; //мю, коэффициент, зависящий от способа закрепления концов вала (в данном частном случае 2)
            double LA; //лямбда, гибкость вала шнека

            F1 = PI * Pow(diam, 2) / 4 * (1 - Pow(AL, 2));
            J1 = PI * Pow(diam, 4) / 64 * (1 - Pow(AL, 4));
            I = Sqrt(J1 / F1);
            //I = d * Sqrt(1 + Pow(AL, 2)) / 4; //после подстановки J и F в I=Sqrt(J1/F1) (преобразованная формула)
            LA = MU * L / I;

            ///////////(7)РАСЧЁТ ВРЕМЕННОГО МОМЕНТА СОПРОТИВЛЕНИЯ КРУЧЕНИЮ WR; НАПРЯЖЕНИЯ КРУЧЕНИЯ TAUmax И РАСПРЕДЕЛЕННОЙ НАГРУЗКИ q
            double WR; //временный момент сопротивления кручению
            double TAUmax; //максимальное напряжение кручения
            double q; //распреленная нагрузка

            WR = PI * Pow(diam, 3) * (1 - Pow(AL, 4)) / 16; //Вывод, м3
            WROutput.Text = $"W_р = {WR} м^3";
            TAUmax = MKR / WR;
            q = RO * gi * L;

            //hardnessPlot.Title
        }

        //Кнопка создания отчёта
        private void ExtrReportCreation(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "PDF Files (*.pdf)|*.pdf";
            saveFileDialog.FileName = "Результат расчётов";

            if (saveFileDialog.ShowDialog() == true)
            {
                string path = saveFileDialog.FileName;

                try
                {
                    Document document = new Document();
                    PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(path, FileMode.Create));
                    document.Open();

                    BaseFont baseFont = BaseFont.CreateFont("D:\\Users\\Garnik\\Desktop\\учёба\\Диплом\\GenShnekApp\\GenShnekApp\\4852-font.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    Font headerFont1 = new Font(baseFont, 16, Font.BOLD);
                    Font headerFont2 = new Font(baseFont, 14, Font.BOLD);
                    Font textFont = new Font(baseFont, 14, Font.NORMAL);

                    iTextSharp.text.Paragraph enter = new iTextSharp.text.Paragraph(" ");

                    iTextSharp.text.Paragraph repHeader = new iTextSharp.text.Paragraph("ОТЧЁТ ПО РАСЧЁТУ ЭКСТРУЗИОННОГО ШНЕКА", headerFont1);
                    repHeader.Alignment = iTextSharp.text.Element.ALIGN_CENTER;

                    iTextSharp.text.Paragraph extrInfoHeader = new iTextSharp.text.Paragraph("Информация о шнеке", headerFont2);
                    extrInfoHeader.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                    iTextSharp.text.Paragraph repMethod = new iTextSharp.text.Paragraph("Способ построения шнека: " + extrMethod + ".", textFont);
                    iTextSharp.text.Paragraph repName = new iTextSharp.text.Paragraph("Наименование шнека: " + extrName + ".", textFont);
                    iTextSharp.text.Paragraph repDiam = new iTextSharp.text.Paragraph($"Диаметр шнека: {extrDiam} мм.", textFont);
                    iTextSharp.text.Paragraph repLength = new iTextSharp.text.Paragraph($"Отношение длины к диаметру: L/D = {extrCoffLength}.", textFont);
                    
                    iTextSharp.text.Paragraph extrOutputHeader = new iTextSharp.text.Paragraph("Проверочные расчёты", headerFont2);
                    extrOutputHeader.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                    iTextSharp.text.Paragraph repQ = new iTextSharp.text.Paragraph($"Производительность: {QOutput.Text}.", textFont);
                    iTextSharp.text.Paragraph repMKR = new iTextSharp.text.Paragraph($"Крутящий момент: {MKROutput.Text}.", textFont);
                    iTextSharp.text.Paragraph repSos = new iTextSharp.text.Paragraph($"Осевое усилие от давления формования: {SosOutput.Text}.", textFont);
                    iTextSharp.text.Paragraph repWR = new iTextSharp.text.Paragraph($"Временный момент сопротивления кручению: {WROutput.Text}.", textFont);

                    iTextSharp.text.Paragraph extrGraph1Header = new iTextSharp.text.Paragraph("Эпюра прочности", headerFont2);
                    extrGraph1Header.Alignment = iTextSharp.text.Element.ALIGN_CENTER;

                    iTextSharp.text.Paragraph extrGraph2Header = new iTextSharp.text.Paragraph("Эпюра жёсткости", headerFont2);
                    extrGraph2Header.Alignment = iTextSharp.text.Element.ALIGN_CENTER;

                    iTextSharp.text.Paragraph extrGraph3Header = new iTextSharp.text.Paragraph("Эпюра устойчивости", headerFont2);
                    extrGraph3Header.Alignment = iTextSharp.text.Element.ALIGN_CENTER;

                    document.Add(repHeader);
                    
                    document.Add(extrInfoHeader);
                    document.Add(repMethod);
                    document.Add(repName);
                    document.Add(repDiam);
                    document.Add(repLength);
                    document.Add(enter);

                    document.Add(extrOutputHeader);
                    document.Add(repQ);
                    document.Add(repMKR);
                    document.Add(repSos);
                    document.Add(repWR);
                    document.Add(enter);

                    document.Add(extrGraph1Header);
                    //document.Add(image1);
                    document.Add(enter);

                    document.Add(extrGraph2Header);
                    document.Add(enter);

                    document.Add(extrGraph3Header);
                    document.Add(enter);

                    document.Close();

                    Process.Start(path);
                }
                catch (IOException ex)
                {
                    MessageBox.Show($"Невозможно создать PDF файл.\n\nДетали:\n{ex}", "File Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"An error occurred.\n\nДетали:\n{ex}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }
}
