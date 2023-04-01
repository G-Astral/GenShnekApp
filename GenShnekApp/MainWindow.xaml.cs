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
        int holeDiamConv;
        int holeDistanceConv;
        int tubeLengthConv;
        int shnekThickConv;
        int shnekDiamConv;
        int hexSizeConv;
        int stepConv;

        double holeDiam;
        double tubeLength;
        double shnekThick;
        double shnekDiam;
        double hexSize;
        double holeDistance;
        double tubeRad;
        double step;

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
            switch (ShnekType.SelectedIndex)
            {
                case 0:
                    if (ImgSketch != null) ImgSketch.Source = (ImageSource)new ImageSourceConverter().ConvertFrom(new Uri(@"D:\Users\Garnik\Desktop\учёба\Диплом\GenShnekApp\GenShnekApp\ShnekSketch1.png"));
                    styleCount = 2;
                    break;
                case 1:
                    if (ImgSketch != null) ImgSketch.Source = (ImageSource)new ImageSourceConverter().ConvertFrom(new Uri(@"D:\Users\Garnik\Desktop\учёба\Диплом\GenShnekApp\GenShnekApp\ShnekSketch2.png"));
                    styleCount = 3;
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
        private void DeleteSpaces1(object sender, TextChangedEventArgs e)
        {
            inputStep.Text = inputStep.Text.Replace(" ", string.Empty);
            inputStep.SelectionStart = inputStep.Text.Length;
        }
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
        private void DeleteSpaces6(object sender, TextChangedEventArgs e)
        {
            inputShnekThick.Text = inputShnekThick.Text.Replace(" ", string.Empty);
            inputShnekThick.SelectionStart = inputShnekThick.Text.Length;
        }

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
            ksDoc3d.fileName = "Шнек";

            part = ksDoc3d.GetPart((int)Part_Type.pTop_Part); // новый компонент
            ksEntity basePlaneXOY = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);  // получим интерфейс базовой плоскости XOY
            ksEntity basePlaneZOY = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeYOZ);  // получим интерфейс базовой плоскости YOZ
            ksEntity basePlaneXOZ = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOZ);  // получим интерфейс базовой плоскости XOZ

            ///////////////////////////Создание кастомных плоскостей/////////////////////////////
            //пока не удаляю, так как в дальнейшем может понадобиться
            /*            ksEntity basePlaneOffsetUP = (ksEntity)part.NewEntity((short)Obj3dType.o3d_planeOffset);
                        PlaneOffsetDefinition offsetPlaneDefUP = basePlaneOffsetUP.GetDefinition();
                        offsetPlaneDefUP.direction = true;
                        offsetPlaneDefUP.offset = -60;
                        offsetPlaneDefUP.SetPlane(basePlaneZOY);
                        basePlaneOffsetUP.Create();*/

            //Выбор исполнения шнека
            switch (ShnekStyle.SelectedIndex)
            {
                case 0:
                    CylinderCreation(tubeRad, tubeLength, basePlaneZOY);
                    JointCreation1(hexSize, holeDistance, basePlaneZOY);
                    HoleCreation(holeDiam, hexSize, basePlaneXOZ, holeDistance, 0);
                    SpyralCreation(tubeRad, step, tubeLength, true, true, shnekThick, shnekDiam, basePlaneZOY, basePlaneXOZ);
                    break;
                case 1:
                    CylinderCreation(tubeRad, tubeLength, basePlaneZOY);
                    SpyralCreation(tubeRad, step, tubeLength, true, true, shnekThick, shnekDiam, basePlaneZOY, basePlaneXOZ);
                    JointCreation2(90, 85, 100, basePlaneZOY);
                    break;
            }
        }

        ///////////////////////////Создание трубы шнека/////////////////////////////
        private void CylinderCreation(double rad, double length, ksEntity plane)
        {
            ksEntity ksSketchE = part.NewEntity((int)Obj3dType.o3d_sketch); // создание нового скетча

            SketchDefinition ksSketchDef = ksSketchE.GetDefinition(); // интерфейс свойств эскиза

            ksSketchDef.SetPlane(plane);  // установим плоскость XOY базовой для эскиза
            ksSketchE.Create();          // создадим эскиз
            ksDocument2D Sketch2D = (ksDocument2D)ksSketchDef.BeginEdit();

            Sketch2D.ksCircle(0, 0, rad, 1);

            ksSketchDef.EndEdit(); // заканчивает редактирование эскиза

            ksEntity bossExtr = part.NewEntity((short)Obj3dType.o3d_baseExtrusion); // сущность для выдавливания
            ksBaseExtrusionDefinition extrDef = bossExtr.GetDefinition(); // интерфейс настроек выдавливания
            ksExtrusionParam extrProp = (ksExtrusionParam)extrDef.ExtrusionParam();

            if (extrProp != null)
            {
                extrDef.SetSketch(ksSketchE); // эскиз операции выдавливания

                extrProp.direction = (short)Direction_Type.dtNormal;      // выбор направления выдавливания
                extrProp.typeNormal = (short)End_Type.etBlind;      // тип выдавливания (строго на глубину)
                extrProp.depthNormal = length;         // глубина выдавливания
                bossExtr.Create();                // создадим операцию
            }
        }

        ///////////////////////////Создание отверстия шестигранника/////////////////////////////
        private void HoleCreation(double diam, double length, ksEntity plane, double x, double y)
        {
            ksEntity ksSketchE = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef = ksSketchE.GetDefinition();

            ksSketchDef.SetPlane(plane);
            ksSketchE.Create();
            ksDocument2D Sketch2D = (ksDocument2D)ksSketchDef.BeginEdit();

            Sketch2D.ksCircle(x, y, diam / 2, 1);

            ksSketchDef.EndEdit();

            ksEntity bossExtr = part.NewEntity((short)Obj3dType.o3d_cutExtrusion);
            ksCutExtrusionDefinition extrDef = bossExtr.GetDefinition();
            ksExtrusionParam extrProp = (ksExtrusionParam)extrDef.ExtrusionParam();

            if (extrProp != null)
            {
                extrDef.SetSketch(ksSketchE);

                extrProp.direction = (short)Direction_Type.dtBoth;
                extrProp.typeReverse = (short)End_Type.etBlind;
                extrProp.depthNormal = length;
                extrProp.depthReverse = length;
                bossExtr.Create();
            }
        }

        ///////////////////////////Создание присоединительного элемента 1/////////////////////////////
        private void JointCreation1(double size, double length, ksEntity plane)
        {
            ksEntity ksSketchE = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef = ksSketchE.GetDefinition();

            ksSketchDef.SetPlane(plane);
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

            ksEntity bossExtr = part.NewEntity((short)Obj3dType.o3d_baseExtrusion);
            ksBaseExtrusionDefinition extrDef = bossExtr.GetDefinition();
            ksExtrusionParam extrProp = (ksExtrusionParam)extrDef.ExtrusionParam();

            if (extrProp != null)
            {
                extrDef.SetSketch(ksSketchE);

                extrProp.direction = (short)Direction_Type.dtReverse;
                extrProp.typeNormal = (short)End_Type.etBlind;
                extrProp.depthReverse = length * 3 / 2;
                bossExtr.Create();
            }
        }

        ///////////////////////////Создание присоединительного элемента 2/////////////////////////////
        private void JointCreation2(double diam, double size, double length, ksEntity plane)
        {
            ksEntity ksSketchE = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef = ksSketchE.GetDefinition();

            ksSketchDef.SetPlane(plane);
            ksSketchE.Create();
            ksDocument2D Sketch2D = (ksDocument2D)ksSketchDef.BeginEdit();
            
            Sketch2D.ksCircle(0, 0, diam / 2, 1);

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
                Sketch2D.ksRegularPolygon(triangle);
            }

            ksSketchDef.EndEdit();


            ksEntity bossExtr = part.NewEntity((short)Obj3dType.o3d_baseExtrusion);
            ksBaseExtrusionDefinition extrDef = bossExtr.GetDefinition();
            ksExtrusionParam extrProp = (ksExtrusionParam)extrDef.ExtrusionParam();

            
            if (extrProp != null)
            {
                extrDef.SetSketch(ksSketchE);

                extrProp.direction = (short)Direction_Type.dtReverse;
                extrProp.typeNormal = (short)End_Type.etBlind;
                extrProp.depthReverse = length;
                bossExtr.Create();
            }
        }

        ///////////////////////////Создание винта/////////////////////////////
        private void SpyralCreation(double rad, double spyralStep, double turn, bool buildDir, bool turnDir, double thick, double sDiam, ksEntity plane, ksEntity profilePlane)
        {
            //траектория
            ksEntity ksSketchE1 = part.NewEntity((short)Obj3dType.o3d_cylindricSpiral);

            CylindricSpiralDefinition ksSketchDef1 = ksSketchE1.GetDefinition();

            ksSketchDef1.SetPlane(plane);

            ksSketchDef1.diam = rad*2;
            ksSketchDef1.buildMode = 0;
            ksSketchDef1.step = spyralStep;
            ksSketchDef1.turn = turn / spyralStep;
            ksSketchDef1.buildDir = buildDir;
            ksSketchDef1.turnDir = turnDir;
            ksSketchE1.hidden = true;

            ksSketchE1.Create();

            //выдавливаемый профиль
            ksEntity ksSketchE2 = part.NewEntity((short)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef2 = ksSketchE2.GetDefinition();

            ksSketchDef2.SetPlane(profilePlane);
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
            holeDiamConv = Convert.ToInt32(inputHoleDiam.Text);
            tubeLengthConv = Convert.ToInt32(inputTubeLength.Text);
            shnekThickConv = Convert.ToInt32(inputShnekThick.Text);
            shnekDiamConv = Convert.ToInt32(inputShnekDiam.Text);
            hexSizeConv = Convert.ToInt32(inputHexSize.Text);
            holeDistanceConv = Convert.ToInt32(inputHoleDistance.Text);
            stepConv = Convert.ToInt32(inputStep.Text);

            if (holeDiamConv == 0)
            {
                holeDiamConv = 24;
                MessageBox.Show("Введён неверный диаметр отверстия!\nПараметру присвоено значение по умолчанию!");
            }
            if (tubeLengthConv < 1000)
            {
                tubeLengthConv = 1000;
                MessageBox.Show("Длина шнека меньше миниального!\nПараметру присвоено минимальное значение!");
            }
            if (tubeLengthConv > 2500)
            {
                tubeLengthConv = 2500;
                MessageBox.Show("Длина шнека больше максимального!\nПараметру присвоено максимальное значение!");
            }
            if (hexSizeConv == 0)
            {
                hexSizeConv = 55;
                MessageBox.Show("Введён неверный размер шестигранника!\nПараметру присвоено значение по умолчанию!");
            }
            if (holeDistanceConv == 0)
            {
                holeDistanceConv = 52;
                MessageBox.Show("Введена неверная толщина винта шнека!\nПараметру присвоено значение по умолчанию!");
            }
            if (shnekDiamConv == 0)
            {
                shnekDiamConv = 135;
                MessageBox.Show("Введён неверный внешний диаметр шнека!\nПараметру присвоено значение по умолчанию!");
            }
            if (hexSizeConv * 1.5 >= shnekDiamConv)
            {
                shnekDiamConv = hexSizeConv * 3;
                MessageBox.Show("Внешний диаметр шнека не может быть меньше или равен внутреннему!\nВнешний диаметр был увеличен!");
            }
            if (holeDiamConv * 2 >= hexSizeConv)
            {
                holeDiamConv = 24;
                hexSizeConv = 55;
                MessageBox.Show("Диаметр отверстия не может быть больше боковой грани шестигранника!\nОбеим параметрам присвоено значение по умолчанию");
            }

            holeDiam = holeDiamConv;
            tubeLength = tubeLengthConv;
            shnekThick = shnekThickConv;
            shnekDiam = shnekDiamConv;
            hexSize = hexSizeConv;
            holeDistance = holeDistanceConv;
            tubeRad = hexSize * 0.75;
            step = stepConv;
        }

        private void InputFieldIsActive(bool isActive)
        {
            inputHoleDiam.IsEnabled = isActive;
            inputTubeLength.IsEnabled = isActive;
            //inputShnekThick.IsEnabled = isActive;
            inputShnekDiam.IsEnabled = isActive;
            inputHexSize.IsEnabled = isActive;
            inputHoleDistance.IsEnabled = isActive;
            //inputStep.IsEnabled = isActive;
        }

        private void GOSTSelection1()
        {
            ShnekType.IsEnabled = true;
            typeCount = 2;
            InputFieldIsActive(false);
            ShnekStyle.IsEnabled = false;
        }

        private void GOSTSelection2()
        {
            InputFieldIsActive(true);
            ShnekType.IsEnabled = false;
            ShnekStyle.IsEnabled = true;
        }
    }
}
