﻿using System;
using System.Collections.Generic;
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

namespace GenShnekApp
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        double holeDiam;
        double tubeLength;
        double shnekThick = 2;
        double shnekDiam;
        double hexSize;
        double hex2Size = 90;
        double holeDistance;
        double tubeRad;
        double step = 100;

        KompasObject kompas;
        ksPart part;

        int typeCount;
        int styleCount;

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
/*                default:
                    typeCount = 2;
                    break;*/
            }

            for (int i = 0; i < typeCount; i++) ShnekType.Items.Add($"Тип {i+1}");
            ShnekType.SelectedIndex = 0;
        }

        //Выбор Типа шнека
        private void ShnekTypeSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ShnekStyle.Items.Clear();
            DefaultShnekChoose.Items.Clear();
            switch (ShnekType.SelectedIndex)
            {
                case 0:
                    if (ImgSketch != null) ImgSketch.Source = (ImageSource)new ImageSourceConverter().ConvertFrom(new Uri(@"D:\Users\Garnik\Desktop\учёба\Диплом\GenShnekApp\GenShnekApp\ShnekSketch1.png"));
                    styleCount = 2;
                    DefaultShnekItems1();
                    break;
                case 1:
                    if (ImgSketch != null) ImgSketch.Source = (ImageSource)new ImageSourceConverter().ConvertFrom(new Uri(@"D:\Users\Garnik\Desktop\учёба\Диплом\GenShnekApp\GenShnekApp\ShnekSketch2.png"));
                    styleCount = 2;
                    DefaultShnekItems2();
                    break;
/*                default:
                    styleCount = 2;
                    break;*/
            }

            for (int i = 0; i < styleCount; i++) ShnekStyle.Items.Add($"Исполнение {i + 1}");
            ShnekStyle.SelectedIndex = 0;
        }

        private void TextBoxInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "0123456789".IndexOf(e.Text) < 0;
        }

        private void DeleteSpaces(object sender, TextChangedEventArgs e)
        {
            inputTubeLength.Text = inputTubeLength.Text.Replace(" ", string.Empty);
            inputTubeLength.SelectionStart = inputTubeLength.Text.Length;
        }
/*        private void DeleteSpaces1(object sender, TextChangedEventArgs e)
        {
            inputStep.Text = inputStep.Text.Replace(" ", string.Empty);
            inputStep.SelectionStart = inputStep.Text.Length;
        }*/
        private void DeleteSpaces2(object sender, TextChangedEventArgs e)
        {
            inputHexSize.Text = inputHexSize.Text.Replace(" ", string.Empty);
            inputHexSize.SelectionStart = inputHexSize.Text.Length;
        }
        private void DeleteSpaces3(object sender, TextChangedEventArgs e)
        {
            inputShnekDiam.Text = inputShnekDiam.Text.Replace(" ", string.Empty);
            inputShnekDiam.SelectionStart = inputShnekDiam.Text.Length;
        }
        private void DeleteSpaces4(object sender, TextChangedEventArgs e)
        {
            inputHoleDiam.Text = inputHoleDiam.Text.Replace(" ", string.Empty);
            inputHoleDiam.SelectionStart = inputHoleDiam.Text.Length;
        }
        private void DeleteSpaces5(object sender, TextChangedEventArgs e)
        {
            inputHoleDistance.Text = inputHoleDistance.Text.Replace(" ", string.Empty);
            inputHoleDistance.SelectionStart = inputHoleDistance.Text.Length;
        }
/*        private void DeleteSpaces6(object sender, TextChangedEventArgs e)
        {
            inputShnekThick.Text = inputShnekThick.Text.Replace(" ", string.Empty);
            inputShnekThick.SelectionStart = inputShnekThick.Text.Length;
        }*/

        private void CreationButton(object sender, RoutedEventArgs e)
        {
            ParamConv();

            try
            {
                kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5");
            }
            catch
            {
                kompas = (KompasObject)Activator.CreateInstance(Type.GetTypeFromProgID("KOMPAS.Application.5"));
            }
            if (kompas == null) return;
            kompas.Visible = true;

            ksDocument3D ksDoc3d = kompas.Document3D(); // создание 3д документа

            ksDoc3d.Create(false, true); // false - видимый режим, true - деталь
            ksDoc3d = kompas.ActiveDocument3D(); // указатель на интерфейс 3д модели 
            ksDoc3d.author = "Garnik";   // указание имени автора

            part = ksDoc3d.GetPart((int)Part_Type.pTop_Part); // новый компонент

            //Шнеки первого типа
            if (ShnekType.SelectedIndex == 0)
            {
                tubeRad = hexSize * 0.75;
                //Дефолтные шнеки
                if (DefaultShnekChoose.IsEnabled == true)
                {
                    switch (DefaultShnekChoose.SelectedIndex)
                    {
                        case 0:
                            CylinderCreation(55 * 0.75, tubeLength);
                            JointCreation1(55, 52);
                            JointHoleCreation(24, 52, 0);
                            SpyralCreation(55 * 0.75, 100, tubeLength, 2, 135);
                            break;
                        case 1:
                            CylinderCreation(55 * 0.75, tubeLength);
                            JointCreation1(55, 52);
                            JointHoleCreation(24, 52, 0);
                            SpyralCreation(55 * 0.75, 100, tubeLength, 2, 150);
                            break;
                        case 2:
                            CylinderCreation(55 * 0.75, tubeLength);
                            JointCreation1(55, 52);
                            JointHoleCreation(24, 52, 0);
                            SpyralCreation(55 * 0.75, 100, tubeLength, 2, 180);
                            break;
                        case 3:
                            CylinderCreation(60 * 0.75, tubeLength);
                            JointCreation1(60, 55);
                            JointHoleCreation(27, 55, 0);
                            SpyralCreation(60 * 0.75, 100, tubeLength, 2, 200);
                            break;
                        case 4:
                            CylinderCreation(60 * 0.75, tubeLength);
                            JointCreation1(60, 55);
                            JointHoleCreation(27, 55, 0);
                            SpyralCreation(60 * 0.75, 100, tubeLength, 2, 300);
                            break;
                        case 5:
                            JointCreation2(90, 95 * 3 / 2);
                            CylinderCreation(90 * 0.75, tubeLength);
                            JointHoleCreation(30, 95, 0);
                            SpyralCreation(90 * 0.75, 100, tubeLength, 2, 300);
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
                            SpyralCreation(tubeRad, step, tubeLength, shnekThick, shnekDiam);
                            break;
                        case 1:
                            JointCreation2(hex2Size, holeDistance * 3 / 2);
                            CylinderCreation(tubeRad, tubeLength);
                            JointHoleCreation(holeDiam, holeDistance, 0);
                            SpyralCreation(tubeRad, step, tubeLength, shnekThick, shnekDiam);
                            break;
                    }
                }
            }
            //Шнеки второго типа
            else
            {
                shnekDiam = 270;
                tubeRad = (shnekDiam * 10 / 27) / 2;
                //Дефолтные шнеки
                if (DefaultShnekChoose.IsEnabled == true)
                {
                    switch (DefaultShnekChoose.SelectedIndex)
                    {
                        case 0:
                            CylinderCreation(tubeRad, tubeLength);
                            SpyralCreation(tubeRad, step, tubeLength, shnekThick, shnekDiam);
                            JointCreation3(90, tubeLength);
                            break;
                        case 1:
                            CylinderCreation(tubeRad, tubeLength);
                            SpyralCreation(tubeRad, step, tubeLength, shnekThick, shnekDiam);
                            break;
                        case 2:
                            CylinderCreation(tubeRad, tubeLength);
                            SpyralCreation(tubeRad, step, tubeLength, shnekThick, shnekDiam);
                            break;
                    }
                }
                //Выбор исполнения шнека
                else
                {
                    if (ShnekStyle.SelectedIndex == 0)
                    {
                        MessageBox.Show("Шнек типа 2 исполнения 1 (ШС-200)\nне предусмотрен для параметризации!");
                    }
                    else
                    {
                        CylinderCreation(tubeRad, tubeLength);
                        SpyralCreation(tubeRad, step, tubeLength, shnekThick, shnekDiam);
                    }
                }
            }
        }

        ///////////////////////////Создание трубы шнека/////////////////////////////
        private void CylinderCreation(double rad, double length)
        {

            ksEntity basePlaneZOY = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeYOZ);  // получим интерфейс базовой плоскости YOZ

            ksEntity ksSketchE = part.NewEntity((int)Obj3dType.o3d_sketch); // создание нового скетча

            SketchDefinition ksSketchDef = ksSketchE.GetDefinition(); // интерфейс свойств эскиза

            ksSketchDef.SetPlane(basePlaneZOY);  // установим плоскость XOY базовой для эскиза
            ksSketchE.Create();          // создадим эскиз
            ksDocument2D Sketch2D = (ksDocument2D)ksSketchDef.BeginEdit();

            Sketch2D.ksCircle(0, 0, rad, 1);

            ksSketchDef.EndEdit(); // заканчивает редактирование эскиза

            ksEntity baseExtr = part.NewEntity((short)Obj3dType.o3d_baseExtrusion); // сущность для выдавливания
            ksBaseExtrusionDefinition extrDef = baseExtr.GetDefinition(); // интерфейс настроек выдавливания
            ksExtrusionParam extrProp = (ksExtrusionParam)extrDef.ExtrusionParam();

            if (extrProp != null)
            {
                extrDef.SetSketch(ksSketchE); // эскиз операции выдавливания

                extrProp.direction = (short)Direction_Type.dtNormal;      // выбор направления выдавливания
                extrProp.typeNormal = (short)End_Type.etBlind;      // тип выдавливания (строго на глубину)
                extrProp.depthNormal = length;         // глубина выдавливания
                baseExtr.Create();                // создадим операцию
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
                hex.radius = size/2;
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
        private void JointCreation3(double diam, double length)
        {
            ksEntity basePlaneZOY = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeYOZ);

            double rad1 = (diam * 0.8) / 2;
            double rad2 = diam / 2;
            double len1 = 174 * 0.05;
            double len2 = 174 * 0.9;
            double len3 = 174 * 0.025;
            double len4 = 174 * 0.025;


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

            Sketch2D2.ksCircle(0, 0, rad2, 1);

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
        private void SpyralCreation(double rad, double spyralStep, double turn, double thick, double sDiam)
        {
            ksEntity basePlaneZOY = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeYOZ);
            ksEntity basePlaneXOZ = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOZ);

            //траектория
            ksEntity ksSketchE1 = part.NewEntity((short)Obj3dType.o3d_cylindricSpiral);

            CylindricSpiralDefinition ksSketchDef1 = ksSketchE1.GetDefinition();

            ksSketchDef1.SetPlane(basePlaneZOY);

            ksSketchDef1.diam = rad*2;
            ksSketchDef1.buildMode = 0;
            ksSketchDef1.step = spyralStep;
            ksSketchDef1.turn = turn / spyralStep;
            ksSketchDef1.buildDir = true;
            ksSketchDef1.turnDir = true;
            ksSketchE1.hidden = true;

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
                rect.x = -thick;
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

        private void ParamConv()
        {
            holeDiam = Convert.ToDouble(inputHoleDiam.Text);
            tubeLength = Convert.ToDouble(inputTubeLength.Text);
            shnekDiam = Convert.ToDouble(inputShnekDiam.Text);
            hexSize = Convert.ToDouble(inputHexSize.Text);
            holeDistance = Convert.ToDouble(inputHoleDistance.Text);

            if (holeDiam == 0)
            {
                holeDiam = 24;
                MessageBox.Show("Введён неверный диаметр отверстия!\nПараметру присвоено значение по умолчанию!");
            }
            if (tubeLength < 1000)
            {
                tubeLength = 1000;
                MessageBox.Show("Длина шнека меньше миниального!\nПараметру присвоено минимальное значение!");
            }
            if (tubeLength > 2500)
            {
                tubeLength = 2500;
                MessageBox.Show("Длина шнека больше максимального!\nПараметру присвоено максимальное значение!");
            }
            if (hexSize == 0)
            {
                hexSize = 55;
                MessageBox.Show("Введён неверный размер шестигранника!\nПараметру присвоено значение по умолчанию!");
            }
            if (holeDistance == 0)
            {
                holeDistance = 52;
                MessageBox.Show("Введена неверная толщина винта шнека!\nПараметру присвоено значение по умолчанию!");
            }
            if (shnekDiam == 0)
            {
                shnekDiam = 135;
                MessageBox.Show("Введён неверный внешний диаметр шнека!\nПараметру присвоено значение по умолчанию!");
            }
            if (hexSize * 1.5 >= shnekDiam)
            {
                shnekDiam = hexSize * 3;
                MessageBox.Show("Внешний диаметр шнека не может быть меньше или равен внутреннему!\nВнешний диаметр был увеличен!");
            }
            if (holeDiam * 2 >= hexSize)
            {
                holeDiam = 24;
                hexSize = 55;
                MessageBox.Show("Диаметр отверстия не может быть больше боковой грани шестигранника!\nОбоим параметрам присвоено значение по умолчанию");
            }
            //tubeRad = hexSize * 0.75;
        }

        private void InputFieldIsActive(bool isActive)
        {
            inputHoleDiam.IsEnabled = isActive;
            //inputTubeLength.IsEnabled = isActive;
            //inputShnekThick.IsEnabled = isActive;
            inputShnekDiam.IsEnabled = isActive;
            inputHexSize.IsEnabled = isActive;
            inputHoleDistance.IsEnabled = isActive;
            //inputStep.IsEnabled = isActive;
        }

        private void GOSTSelection1()
        {
            //ShnekType.IsEnabled = true;
            typeCount = 2;
            InputFieldIsActive(false);
            ShnekStyle.IsEnabled = false;
            DefaultShnekChoose.IsEnabled = true;
        }

        private void GOSTSelection2()
        {
            InputFieldIsActive(true);
            //ShnekType.IsEnabled = false;
            ShnekStyle.IsEnabled = true;
            DefaultShnekChoose.IsEnabled = false;
        }

        private void DefaultShnekItems1()
        {
            for (int i = 0; i < 6; i++)
            {
                if (i == 0)
                {
                    DefaultShnekChoose.Items.Add($"ШБ-135");
                }
                if (i == 1)
                {
                    DefaultShnekChoose.Items.Add($"ШБ-150");
                }
                if (i == 2)
                {
                    DefaultShnekChoose.Items.Add($"ШБ-180");
                }
                if (i == 3)
                {
                    DefaultShnekChoose.Items.Add($"ШБ-200");
                }
                if (i == 4)
                {
                    DefaultShnekChoose.Items.Add($"ШБ-300");
                }
                if (i == 5)
                {
                    DefaultShnekChoose.Items.Add($"ШБ-300У");
                }
            }
            DefaultShnekChoose.SelectedIndex = 0;
        }
        private void DefaultShnekItems2()
        {
            for (int i = 0; i < 3; i++)
            {
                if (i == 0)
                {
                    DefaultShnekChoose.Items.Add($"ШС-80");
                }
                if (i == 1)
                {
                    DefaultShnekChoose.Items.Add($"ШС-100");
                }
                if (i == 2)
                {
                    DefaultShnekChoose.Items.Add($"ШС-200");
                }
            }
            DefaultShnekChoose.SelectedIndex = 0;
        }
    }
}
