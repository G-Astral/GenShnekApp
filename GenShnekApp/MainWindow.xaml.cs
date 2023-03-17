using System;
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
        
        int typeCount;
        int styleCount;

        public MainWindow()
        {
            InitializeComponent();

        }

        private void GhostTypeSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ShnekType.Items.Clear();
            switch (GhostType.SelectedIndex)
            {
                case 0:
                    ShnekType.IsEnabled = true;
                    typeCount = 2;

                    holeDiamConv = Convert.ToInt32(inputHoleDiam.Text);
                    tubeLengthConv = Convert.ToInt32(inputTubeLength.Text);
                    shnekThickConv = Convert.ToInt32(inputShnekThick.Text);
                    shnekDiamConv = Convert.ToInt32(inputShnekDiam.Text);
                    hexSizeConv = Convert.ToInt32(inputHexSize.Text);
                    holeDistanceConv = Convert.ToInt32(inputHoleDistance.Text);
                    stepConv = Convert.ToInt32(inputStep.Text);

                    holeDiam = holeDiamConv;
                    tubeLength = tubeLengthConv;
                    shnekThick = shnekThickConv;
                    shnekDiam = shnekDiamConv;
                    hexSize = hexSizeConv;
                    holeDistance = holeDistanceConv;
                    tubeRad = hexSize * 0.75;
                    step = stepConv;
                    break;
                case 1:
                    ShnekType.IsEnabled = false;
                    break;
                default:
                    typeCount = 2;
                    break;
            }

            for (int i = 0; i < typeCount; i++) ShnekType.Items.Add($"Тип {i+1}");
            ShnekType.SelectedIndex = 0;
        }

        private void ShnekTypeSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            switch (ShnekType.SelectedIndex)
            {
                case 0:
                    if (ImgSketch != null) ImgSketch.Source = (ImageSource)new ImageSourceConverter().ConvertFrom(new Uri(@"D:\Users\Garnik\Desktop\учёба\Диплом\GenShnekApp\GenShnekApp\ShnekSketch1.png"));
                    styleCount = 2;
                    break;
                case 1:
                    if (ImgSketch != null) ImgSketch.Source = (ImageSource)new ImageSourceConverter().ConvertFrom(new Uri(@"D:\Users\Garnik\Desktop\учёба\Диплом\GenShnekApp\GenShnekApp\ShnekSketch2.png"));
                    styleCount = 2;
                    break;
                default:
                    break;
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

            if (holeDiamConv == 0)
            {
                holeDiamConv = 24;
                MessageBox.Show("Введён неверный диаметр отверстия!\nПараметру присвоено значение по умолчанию!");
            }
            if (tubeLengthConv < 1000)
            {
                tubeLengthConv = 1000;
                MessageBox.Show("Введена неверная длина шнека!" +
                    "\nПараметру присвоено значение по умолчанию!");
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
            if (hexSizeConv*1.5 >= shnekDiamConv)
            {
                shnekDiamConv = hexSizeConv*3;
                MessageBox.Show("Внешний диаметр шнека не может быть меньше или равен внутреннему!\nВнешний диаметр был увеличен!");
            }
            if (holeDiamConv*2 >= hexSizeConv)
            {
                holeDiamConv = 24;
                hexSizeConv = 55;
                MessageBox.Show("Диаметр отверстия не может быть больше боковой грани шестигранника!\nОбеим параметрам присвоено значение по умолчанию");
            }



            //KompasObject kompas;
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

            ksPart part = ksDoc3d.GetPart((int)Part_Type.pTop_Part); // новый компонент
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

            switch (ShnekStyle.SelectedIndex)
            {
                case 0:
                    ///////////////////////////Создание трубы шнека/////////////////////////////
                    CylinderCreation(tubeRad, tubeLength, basePlaneZOY, 0, 0, (part)part);

                    ///////////////////////////Создание шестигранника/////////////////////////////
                    HexCreation(hexSize, holeDistance, basePlaneZOY, (part)part);

                    ///////////////////////////Создание отверстия шестигранника/////////////////////////////
                    HoleCreation(holeDiam, hexSize, basePlaneXOZ, holeDistance, 0, (part)part);

                    ///////////////////////////Создание винта/////////////////////////////
                    SpyralCreation(tubeRad, step, tubeLength, true, true, shnekThick, shnekDiam, basePlaneZOY, basePlaneXOZ, (part)part);
                    break;
                case 1:
                    ///////////////////////////Создание трубы шнека/////////////////////////////
                    CylinderCreation(tubeRad, tubeLength, basePlaneZOY, 0, 0, (part)part);

                    ///////////////////////////Создание шестигранника/////////////////////////////
                    HexCreation(hexSize, holeDistance, basePlaneZOY, (part)part);

                    ///////////////////////////Создание отверстия шестигранника/////////////////////////////
                    HoleCreation(holeDiam, hexSize, basePlaneXOZ, holeDistance, 0, (part)part);

                    break;
                default:
                    break;
            }

    

        }

        private void CylinderCreation(double rad, double length, ksEntity plane, double x, double y, part part)
        {
            ksEntity ksSketchE = part.NewEntity((int)Obj3dType.o3d_sketch); // создание нового скетча

            SketchDefinition ksSketchDef = ksSketchE.GetDefinition(); // интерфейс свойств эскиза

            ksSketchDef.SetPlane(plane);  // установим плоскость XOY базовой для эскиза
            ksSketchE.Create();          // создадим эскиз
            ksDocument2D Sketch2D = (ksDocument2D)ksSketchDef.BeginEdit();

            Sketch2D.ksCircle(x, y, rad, 1);

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

        private void HoleCreation(double diam, double length, ksEntity plane, double x, double y, part part)
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

        private void HexCreation(double size, double length, ksEntity plane, part part)
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

        private void SpyralCreation(double rad, double spyralStep, double turn, bool buildDir, bool turnDir, double thick, double sDiam, ksEntity plane, ksEntity profilePlane, part part)
        {
            //траектория
            ksEntity ksSketchE3 = part.NewEntity((short)Obj3dType.o3d_cylindricSpiral);

            CylindricSpiralDefinition ksSketchDef3 = ksSketchE3.GetDefinition();

            ksSketchDef3.SetPlane(plane);

            ksSketchDef3.diam = rad*2;
            ksSketchDef3.buildMode = 0;
            ksSketchDef3.step = spyralStep;
            ksSketchDef3.turn = turn / spyralStep;
            ksSketchDef3.buildDir = buildDir;
            ksSketchDef3.turnDir = turnDir;
            ksSketchE3.hidden = true;

            ksSketchE3.Create();

            //выдавливаемый профиль
            ksEntity ksSketchE4 = part.NewEntity((short)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef4 = ksSketchE4.GetDefinition();

            ksSketchDef4.SetPlane(profilePlane);
            ksSketchE4.hidden = true;
            ksSketchE4.Create();
            ksDocument2D Sketch2D4 = (ksDocument2D)ksSketchDef4.BeginEdit();

            ksRectangleParam rect2 = (ksRectangleParam)kompas.GetParamStruct((short)StructType2DEnum.ko_RectangleParam);
            if (rect2 != null)
            {
                // Параметры прямоугольника
                rect2.ang = 0;
                rect2.x = -thick;
                rect2.y = rad;
                rect2.width = thick;
                rect2.height = sDiam / 2 - rad;
                rect2.style = 1;
                Sketch2D4.ksRectangle(rect2);
            }

            ksSketchDef4.EndEdit();

            //выдавливание профиля по траектории
            ksEntity trajectoryExtr5 = part.NewEntity((short)Obj3dType.o3d_baseEvolution);
            ksBaseEvolutionDefinition extrDef5 = trajectoryExtr5.GetDefinition();

            extrDef5.PathPartArray().add(ksSketchE3);
            extrDef5.SetSketch(ksSketchE4);
            trajectoryExtr5.Create();
        }

    }
}
