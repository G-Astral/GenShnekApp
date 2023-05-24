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
using static System.Math;

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

        double extrDiam;
        double extrRad;
        double extrCoffLength;
        double extrLength;

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
            }
        }
        private void DefaultShnekChooseSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (GhostType.SelectedIndex == 2)
            {
                switch (DefaultShnekChoose.SelectedIndex)
                {
                    case 0:
                        inputExtrShnekDiam.Text = "20";
                        inputExtrShnekCoffLength.Text = "20";
                        break;
                    case 1:
                        inputExtrShnekDiam.Text = "32";
                        inputExtrShnekCoffLength.Text = "20";
                        break;
                    case 2:
                        inputExtrShnekDiam.Text = "45";
                        inputExtrShnekCoffLength.Text = "20";
                        break;
                    case 3:
                        inputExtrShnekDiam.Text = "45";
                        inputExtrShnekCoffLength.Text = "25";
                        break;
                    case 4:
                        inputExtrShnekDiam.Text = "63";
                        inputExtrShnekCoffLength.Text = "20";
                        break;
                    case 5:
                        inputExtrShnekDiam.Text = "63";
                        inputExtrShnekCoffLength.Text = "25";
                        break;
                    case 6:
                        inputExtrShnekDiam.Text = "63";
                        inputExtrShnekCoffLength.Text = "30";
                        break;
                    case 7:
                        inputExtrShnekDiam.Text = "90";
                        inputExtrShnekCoffLength.Text = "20";
                        break;
                    case 8:
                        inputExtrShnekDiam.Text = "90";
                        inputExtrShnekCoffLength.Text = "25";
                        break;
                    case 9:
                        inputExtrShnekDiam.Text = "90";
                        inputExtrShnekCoffLength.Text = "30";
                        break;
                    case 10:
                        inputExtrShnekDiam.Text = "125";
                        inputExtrShnekCoffLength.Text = "25";
                        break;
                    case 11:
                        inputExtrShnekDiam.Text = "160";
                        inputExtrShnekCoffLength.Text = "20";
                        break;
                    case 12:
                        inputExtrShnekDiam.Text = "200";
                        inputExtrShnekCoffLength.Text = "20";
                        break;
                }
            }
        }

        private void TextBoxInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "0123456789".IndexOf(e.Text) < 0;
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
                    //Дефолтные шнеки
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
                    //Дефолтные шнеки
                    if (DefaultShnekChoose.IsEnabled == true)
                    {
                        switch (DefaultShnekChoose.SelectedIndex)
                        {
                            case 0:
                                shnekDiam = 80;
                                tubeRad = (shnekDiam * 10 / 18) / 2;
                                CylinderCreation(tubeRad, tubeLength);
                                SpyralCreation(tubeRad, step, 0, tubeLength, shnekThick, shnekDiam);
                                JointCreation4(tubeRad * 2, 56);
                                break;
                            case 1:
                                shnekDiam = 100;
                                tubeRad = (shnekDiam * 10 / 18) / 2;
                                CylinderCreation(tubeRad, tubeLength);
                                SpyralCreation(tubeRad, step, 0, tubeLength, shnekThick, shnekDiam);
                                JointCreation4(tubeRad * 2, 174);
                                break;
                            case 2:
                                MessageBox.Show("Отверстие шнека типа 2 исполнения 1 на данный момент не реализовано");
                                shnekDiam = 200;
                                tubeRad = (shnekDiam * 10 / 18) / 2;
                                CylinderCreation(tubeRad, tubeLength);
                                SpyralCreation(tubeRad, step, 0, tubeLength, shnekThick, shnekDiam);
                                JointCreation3(tubeRad * 2 * 0.9, tubeLength);
                                break;
                        }
                    }
                    //Выбор исполнения шнека
                    else
                    {
                        if (ShnekStyle.SelectedIndex == 0)
                        {
                            MessageBox.Show("Отверстие шнека типа 2 исполнения 1 на данный момент не реализовано");
                            CylinderCreation(tubeRad, tubeLength);
                            SpyralCreation(tubeRad, step, 0, tubeLength, shnekThick, shnekDiam);
                            HoleType2Creation1(tubeRad);
                        }
                        else
                        {
                            MessageBox.Show("Шнек типа 2 исполнения 2 на данный момент не реализован");
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
                    extrRad = extrDiam / 2;
                    CylinderCreation(extrRad, extrLength);
                    SpyralCreation(extrRad * 1.2, extrDiam * 1.2, 0, extrLength, extrDiam * 0.06, extrDiam);
                    ConeCreation(extrRad);
                    ShnekCalc(extrDiam, extrLength);
                }
                else
                {
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
        private void JointCreation3(double diam, double length)
        {
            ksEntity basePlaneZOY = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeYOZ);

            double rad1 = (diam * 0.8) / 2;
            double rad2 = diam / 2;
            double len1 = 174 * 0.1;
            double len2 = 174 * 0.8;
            double len3 = 174 * 0.05;
            double len4 = 174 * 0.05;

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

        ///////////////////////////Создание присоединительного элемента 4 (тип 2 исполнение 2)/////////////////////////////
        private void JointCreation4(double diam, double length)
        {
            ksEntity basePlaneZOY = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeYOZ);

            length = length * 4 / 3;
            double rad1 = (diam * 0.85) / 2;
            double rad2 = (diam * 0.75) / 2;
            double len1 = length * 0.1;
            double len2 = length * 0.1;
            double len3 = length * 0.75;
            double len4 = length * 0.05;

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

            ksEntity plane3 = OffsetPlaneCreation(-len1 - len2, basePlaneZOY);
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
                bossExtr3.Create();
            }


            ksEntity plane4 = OffsetPlaneCreation(- len1 - len2 - len3, basePlaneZOY);
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

                extrProp4.direction = (short)Direction_Type.dtReverse;
                extrProp4.typeReverse = (short)End_Type.etBlind;
                extrProp4.depthReverse = len4;
                extrProp4.draftOutwardReverse = true;
                extrProp4.draftValueReverse = 45;
                bossExtr4.Create();
            }
        }

        //TODO
        ///////////////////////////Создание сквозного отверстия 1 (тип 2 исполнение 1)/////////////////////////////
        private void HoleType2Creation1(double diam)
        {
            ksEntity basePlaneZOY = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeYOZ);

            double rad = diam / 2;
            double len1 = 163;

            ksEntity ksSketchE1 = part.NewEntity((int)Obj3dType.o3d_sketch);

            SketchDefinition ksSketchDef1 = ksSketchE1.GetDefinition();

            ksSketchDef1.SetPlane(basePlaneZOY);
            ksSketchE1.Create();
            ksDocument2D Sketch2D1 = (ksDocument2D)ksSketchDef1.BeginEdit();

            Sketch2D1.ksCircle(0, 0, rad, 1);
            
            ksRectangleParam rect1 = (ksRectangleParam)kompas.GetParamStruct((short)StructType2DEnum.ko_RectangleParam);
/*            if (rect1 != null)
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

                extrProp1.direction = (short)Direction_Type.dtNormal;
                extrProp1.typeNormal = (short)End_Type.etBlind;
                extrProp1.depthNormal = len1;
                cutExtr1.Create();
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
        private void SpyralCreation(double rad, double spyralStep, double start, double end, double thick, double sDiam)
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
            ksSketchDef1.turn = end / spyralStep;
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
                    textBox.BorderBrush = new SolidColorBrush(Color.FromRgb(171, 173, 179));
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

                if (GhostType.SelectedIndex != 0)
                {
                    if (GhostType.SelectedIndex == 1)
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
                if (i == 0 ) ShnekType.Items.Add("Для пластических масс");
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

        //Расчёт шнека на прочность, жёсткость и устойчивость
        ///////////(1)НАЧАЛО
        private void ShnekCalc(double d, double L)
        {
            //РАСЧЁТ НА ПРОЧНОСТЬ КОНСОЛЬНО ЗАКРЕПЛЁННОГО ШНЕКА
            const double PI = Math.PI; //Число ПИ
            const double g = 9.81; //ускорение свободного падения

            double A; //постоянная прямого потока
            double B; //постоянная обратного потока
            double G; //поток утечки
            double ZX; //число отрезков разбиения
            double Q; //производительность
            double MKR; //крутящий момент
            double F; //площадь поперечного сечения
            double Sos; //осевое усилие от давления формования
            double K; //параметр

            //ОТДЕЛЬНАЯ КАСТА ПЕРЕМЕННЫХ, КОТОРЫЕ ЧЁРТ ЗНАЕТ ОТКУДА ПОЯВИЛИСЬ (ПОКА НЕ ВЫЯСНЮ ОТКУДА, БУДУТ РАВНЫ 1)            
            double DL;
            DL = 1;

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
            double d1 = 0.001; //диаметр осевого отверстия шнека, м
            //double d = 0.032; //наружный диаметр шнека, м
            //d = 0.032; //наружный диаметр шнека, м
            d /= 1000;
            //double L = 0.64; //длина нарезной части шнека, м
            L /= 1000;
            double H = 0.0032; //глубина винтового канала шнека, м
            double FI = 17; //угол наклона винтовой линии шнека, град
            double T = 0.032; //шаг винтовой нарезки шнека, м
            double E1 = 0.0032; //ширина гребня винтового канала шнека, м
            double N = 5; //технологическая мощность, кВт
            double W = 70; //частота вращения шнека, об/мин
            double GAM; //удельный вес материала
            //}

            ///////////(3)ОПРЕДЕЛЕНЕИ ЧИСЛА ОТРЕЗКОВ РАЗБИЕНИЯ ДЛИНЫ НАРЕЗНОЙ ЧАСТИ ШНЕКА
            ZX = L / dX;

            ///////////(4)ОПРЕДЕЛЕНИЕ ПОСТОЯННЫХ ПРЯМОГО, ОБРАТНОГО ПОТОКА, ПОТОКА УТЕЧКИ И ПРОИЗВОДИТЕЛЬНОСТИ
            FI = FI * PI / 180;
            A = PI * d * H * (T - E) * Pow(Cos(FI), 2) / 2;
            B = Pow(H, 3) * (T - E) * Sin(2 * FI) / (24 * L);
            G = Pow(PI, 2) * Pow(d, 2) * Pow(DL, 3) * Tan(FI) * Sin(FI) / (10 * E1 * L);
            //K = Sqrt(P/(E*J)); TODO  найти J по формулам
            K = 1;
            Q = A * K * N / (K - B - G); //Вывод
            QOutput.Text = $"Q = {Q} м^3/с";

            ///////////(5)ОПРЕДЕЛЕНИЕ КРУТЯЩЕГО МОМЕНТА, ПЛОЩАДИ ПОПЕРЕЧНОГО СЕЧЕНИЯ ШНЕКА И ОСЕВОГО УСИЛИЯ
            MKR = 9550 * N / W; //Вывод
            MKROutput.Text = $"M_кр = {MKR} Н*м";
            F = PI * Pow(d, 2) / 4;
            Sos = F * P; //Вывод
            SosOutput.Text = $"S_ос = {Sos} Н";

            ///////////(6)РАСЧЁТ ГИБКОСТИ ШНЕКА
            double AL; //альфа, отношение диаметра осевого отверстия шнека к наружному диаметру шнека
            double F1; // площадь поперечного сечения шнека сечения А-А
            double J1; // момент инерции поперечного сечения А-А
            double I; // радиус инерции сечения
            double MU = 2; //мю, коэффициент, зависящий от способа закрепления концов вала (в данном частном случае 2)
            double LA; //лямбда, гибкость вала шнека

            AL = d1 / d;
            F1 = PI * Pow(d, 2) / 4 * (1 - Pow(AL, 2));
            J1 = PI * Pow(d, 4) / 64 * (1 - Pow(AL, 4));
            I = Sqrt(J1 / F1);
            //I = d * Sqrt(1 + Pow(AL, 2)) / 4; //после подстановки J и F в I=Sqrt(J1/F1) (преобразованная формула)
            LA = MU * L / I;

            ///////////(7)РАСЧЁТ ВРЕМЕННОГО МОМЕНТА СОПРОТИВЛЕНИЯ КРУЧЕНИЮ WR; НАПРЯЖЕНИЯ КРУЧЕНИЯ TAUmax И РАСПРЕДЕЛЕННОЙ НАГРУЗКИ q
            double WR; //временный момент сопротивления кручению
            double TAUmax; //максимальное напряжение кручения
            double q; //распреленная нагрузка

            WR = PI * Pow(d, 3) * (1 - Pow(AL, 4)) / 16; //Вывод, м3
            WROutput.Text = $"W_р = {WR} м^3";
            TAUmax = MKR / WR;
            q = RO * g * L;


        }

        private void UpdateGraphics(object sender, RoutedEventArgs e)
        {
/*            extrDiam = Convert.ToDouble(inputExtrShnekDiam.Text);
            extrCoffLength = Convert.ToDouble(inputExtrShnekCoffLength.Text);
            extrLength = extrDiam * extrCoffLength;
            ShnekCalc(extrDiam, extrLength);*/
            MessageBox.Show("Это кнопка обновления графиков.");
        }

        private void NoteButton(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Расчёты проводятся только для экструзионных шнеков.");
        }
    }
}
