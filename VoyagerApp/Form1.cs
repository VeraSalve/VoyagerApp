using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using System.Device.Location;
using System.Net;
using Excel = Microsoft.Office.Interop.Excel;


// Библиотеки для карты
using GMap.NET;
using GMap.NET.MapProviders;
using GMap.NET.WindowsForms;
using GMap.NET.WindowsForms.Markers;
using GMap.NET.WindowsForms.ToolTips;
using System.Drawing.Imaging;

namespace VoyagerApp
{
    public partial class Form1 : Form
    {

        //string[,] ImportExelList = new string[1000, 100];
        public Form1()
        {
            InitializeComponent();
            GMapProvider.Language = LanguageType.Russian;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ImpExcel();
        }

        private int ImpExcel()
        {
            // Выбрать путь и имя файла в диалоговом окне
            OpenFileDialog ofd = new OpenFileDialog();
            // Задаем расширение имени файла по умолчанию (открывается папка с программой)
            ofd.DefaultExt = "*.xls;*.xlsx";
            // Задаем строку фильтра имен файлов, которая определяет варианты
            ofd.Filter = "файл Excel (Spisok.xlsx)|*.xlsx";
            // Задаем заголовок диалогового окна
            ofd.Title = "Выберите файл базы данных";
            if (!(ofd.ShowDialog() == DialogResult.OK)) // если файл БД не выбран -> Выход
                return 0;
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1-й лист
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//последнюю ячейку
                                                                                                // размеры базы
            int lastColumn = (int)lastCell.Column;
            int lastRow = (int)lastCell.Row;


            //////////////

            ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            ObjWorkExcel.Quit(); // выйти из Excel
            GC.Collect(); // убрать за собой
            label1.Text = ofd.FileName;
            return lastRow;
        }


        static List<InputData> GetAbonentsWithSameAddress(List<InputData> data, decimal minDz, int takeCount)
        {
            var result = new List<InputData>();

            //абоненты которых обходят в первую очередь
            var targets = data
               .Where(x => x.DZ > minDz)
               .OrderByDescending(x => x.DZ)
               .Take(takeCount);

            foreach (var target in targets)
            {
                //абоненты с тем же адресом и дз больше нуля
                var anotherAbonents = data.Where(x => x.Address == target.Address && x.DZ > 0);

                foreach (var abonent in anotherAbonents)
                    if (!result.Contains(abonent))
                        result.Add(abonent);
            }

            return result
                .OrderByDescending(x => x.DZ)
                // .Take(takeCount)
                .ToList();
        }

        public class Cpoint
        {
            public double x { get; set; }
            public double y { get; set; }

            public Cpoint() { }
            public Cpoint(double _x, double _y)
            {
                x = _x;
                y = _y;
            }
        }
        //Слои 
        GMapOverlay PositionsForUser = new GMapOverlay("PositionsForUser");
        GMapOverlay PositionForCicle = new GMapOverlay("PositionGForCicle");
        GMapOverlay PositionForSector = new GMapOverlay("PositionGForSector");
        PointLatLng CentralCoordinate = new PointLatLng(0, 0);
        List<PointLatLng> CircleCoords = new List<PointLatLng>();
        double CircleMiddlePosition_X = 0;
        double CircleMiddlePosition_Y = 0;

        private void gMapControl1_Load(object sender, EventArgs e)
        {
            // Настройки для компонента GMap
            gmap.Bearing = 0;
            // Перетаскивание правой кнопки мыши
            gmap.CanDragMap = true;
            // Перетаскивание карты левой кнопкой мыши
            gmap.DragButton = MouseButtons.Left;

            gmap.GrayScaleMode = true;

            // Все маркеры будут показаны
            gmap.MarkersEnabled = true;
            // Максимальное приближение
            gmap.MaxZoom = 18;
            // Минимальное приближение
            gmap.MinZoom = 2;
            // Курсор мыши в центр карты
            gmap.MouseWheelZoomType = GMap.NET.MouseWheelZoomType.MousePositionWithoutCenter;

            // Отключение негативного режима
            gmap.NegativeMode = false;
            // Разрешение полигонов
            gmap.PolygonsEnabled = true;
            // Разрешение маршрутов
            gmap.RoutesEnabled = true;
            // Скрытие внешней сетки карты
            gmap.ShowTileGridLines = false;
            // При загрузке 10-кратное увеличение
            gmap.Zoom = 10;

            // Чья карта используется
            gmap.MapProvider = GMap.NET.MapProviders.GMapProviders.GoogleMap;

            // Загрузка этой точки на карте
            GMap.NET.GMaps.Instance.Mode = GMap.NET.AccessMode.ServerOnly;
            gmap.Position = new GMap.NET.PointLatLng(51.769476, 55.122400);

            //// Создаём новый список маркеров
            GMapOverlay markersOverlay = new GMapOverlay("markers");

            //// Инициализация красного маркера с указанием его коордиант
            //GMarkerGoogle marker = new GMarkerGoogle(new PointLatLng(47.200041, 38.8855908), GMarkerGoogleType.red);
            //marker.ToolTip = new GMap.NET.WindowsForms.ToolTips.GMapRoundedToolTip(marker);

            //// Текст отображаемый с маркером
            //marker.ToolTipText = "Software Technologies";
            //// Добавляем маркер в список маркеров
            //markersOverlay.Markers.Add(marker);
            //gmap.Overlays.Add(markersOverlay);




        }
        public static PointLatLng FindPointDistanceFrom(PointLatLng startPoint, double initialBearingRadians, double distanceKilometres)
        {

            const double radiusEarthKilometres = 6371.01;
            var distRatio = distanceKilometres / radiusEarthKilometres;
            var distRatioSine = Math.Sin(distRatio);
            var distRatioCosine = Math.Cos(distRatio);

            var startLatRad = DegreesToRadians(startPoint.Lat);
            var startLonRad = DegreesToRadians(startPoint.Lng);

            var startLatCos = Math.Cos(startLatRad);
            var startLatSin = Math.Sin(startLatRad);

            var endLatRads = Math.Asin((startLatSin * distRatioCosine) + (startLatCos * distRatioSine * Math.Cos(initialBearingRadians)));
            var endLonRads = startLonRad + Math.Atan2(Math.Sin(initialBearingRadians) * distRatioSine * startLatCos, distRatioCosine - startLatSin * Math.Sin(endLatRads));
            return new PointLatLng(RadiansToDegrees(endLatRads), RadiansToDegrees(endLonRads));



        }
        public static double DegreesToRadians(double degrees)
        {
            const double degToRadFactor = Math.PI / 180;
            return degrees * degToRadFactor;
        }
        public static double RadiansToDegrees(double radians)
        {
            const double radToDegFactor = 180 / Math.PI;
            return radians * radToDegFactor;
        }
        private void CreateCircle(double lat, double lon, double radius, int ColorIndex)
        {
            gmap.Overlays.Remove(PositionForCicle);
            PointLatLng point = new PointLatLng(lat, lon);
            int segments = 360;

            List<PointLatLng> gpollist = new List<PointLatLng>();
            CircleCoords.Clear();
            for (int i = 0; i < segments; i++)
            {
                gpollist.Add(FindPointDistanceFrom(point, i * (Math.PI / 180), radius / 1000));
            }
            //лист с координатами для дальнейшей отрисовки сектора
            CircleCoords = gpollist.ToList();
            GMapPolygon polygon = new GMapPolygon(gpollist, "Circle");
            polygon.Fill = new SolidBrush(Color.FromArgb(30, Color.Aqua));
            polygon.Stroke = new Pen(Color.Red, 3);
            PositionForCicle.Polygons.Add(polygon);
            gmap.Overlays.Add(PositionForCicle);
        }
        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            //чтобы окружность не показывалась, когда изменяется радиус
            gmap.Overlays.Remove(PositionForCicle);
            PositionForCicle.Clear();
            gmap.Overlays.Add(PositionForCicle);
            CreateCircle(CircleMiddlePosition_Y, CircleMiddlePosition_X, Convert.ToDouble(numericUpDown1.Value), 1);
        }
        public string ChooseFolder()
        {
            string path = Environment.CurrentDirectory;
            folderBrowserDialog1.Description = "Выберите путь для сохранения файла";

            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                path = folderBrowserDialog1.SelectedPath;
                
            }
            return path;
        }
        GMapOverlay OverlayForAddress = new GMapOverlay("Address");
        private void button2_Click(object sender, EventArgs e)
        {

            List<InputData> inpolygonResults = new List<InputData>();
            //string myAPI = "AIzaSyBUfuW1DgZydpfpXDxYxnkGMsJWO5ZP0JM";
            string myAPI = "AIzaSyAD9VnXyK4mmnw7X1J6DZyfDHm9CVDo7Fk";
            // Запрос + считанный из апи ключ
            string zapros = "https://maps.googleapis.com/maps/api/geocode/xml?address={0}&sensor=true_or_false&language=ru&key=" + myAPI;

            if (label1.Text != "Адрес файла")
            {
                var data = InputData.Parse(label1.Text);
                var results = GetAbonentsWithSameAddress(data, whatMinDz.Value, 100);
               
                foreach (var result in results)
                {
                    // Запрос к API
                    string url = string.Format(zapros, Uri.EscapeDataString("Оренбург " + result.Address));

                    // Выполняем запрос
                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);

                    // Получаем ответ от интернет-ресурса
                    WebResponse response = request.GetResponse();

                    // Чтение данных из интернет-ресурса
                    Stream dataStream = response.GetResponseStream();

                    // Чтение
                    StreamReader sreader = new StreamReader(dataStream);

                    // Считывает поток от текущего положения до конца         
                    string responsereader = sreader.ReadToEnd();

                    // Закрываем поток ответа
                    response.Close();

                    // Блок парсинга данных
                    XmlDocument xmldoc = new XmlDocument();
                    xmldoc.LoadXml(responsereader);

                    if (xmldoc.GetElementsByTagName("status")[0].ChildNodes[0].InnerText == "OK")
                    {
                        // Получение широты и долготы
                        XmlNodeList nodes = xmldoc.SelectNodes("//location");
                        double latitude = 0.0;
                        double longitude = 0.0;
                        foreach (XmlNode node in nodes)
                        {
                            latitude = XmlConvert.ToDouble(node.SelectSingleNode("lat").InnerText.ToString());
                            longitude = XmlConvert.ToDouble(node.SelectSingleNode("lng").InnerText.ToString());
                        }

                        // Строка со всеми данными
                        string formatted_address = xmldoc.SelectNodes("//formatted_address").Item(0).InnerText.ToString();
                        // Парсинг №1
                        string[] words = formatted_address.Split(',');
                        string dataMarker = string.Empty;
                        foreach (string word in words)
                        {
                            dataMarker += word + ";" + Environment.NewLine;
                        }

                        double x = latitude;
                        double y = longitude;
                        PointLatLng pointsearch = new PointLatLng(x, y);
                        // Идем по слоям
                        var overlays = gmap.Overlays;
                        for (int i = 0; i < overlays.Count; i++)
                        {
                            // Идем по полигонам
                            var polygons = overlays[i].Polygons;
                            for (int j = 0; j < polygons.Count; j++)
                            {
                                // В каждом полигоне каждого слоя проверяем принадлежность точки полигону
                                if (polygons[j].IsInside(pointsearch))
                                {
                                    GMarkerGoogle pointinarea = new GMarkerGoogle(new PointLatLng(x, y), GMarkerGoogleType.green_big_go);
                                    pointinarea.ToolTip = new GMapRoundedToolTip(pointinarea);
                                    gmap.Overlays.Add(OverlayForAddress);
                                    GMarkerGoogle addressmarker = new GMarkerGoogle(new PointLatLng(latitude, longitude), GMarkerGoogleType.orange);
                                    addressmarker.ToolTip = new GMapRoundedToolTip(addressmarker);
                                    addressmarker.ToolTipMode = MarkerTooltipMode.Always;
                                    addressmarker.ToolTipText = dataMarker;
                                    OverlayForAddress.Markers.Add(addressmarker);
                                    gmap.Position = new PointLatLng(latitude, longitude);

                                    inpolygonResults.Add(result);

                                }
                            }
                        }
                    }
                }
                
                PrintExel.ExportToExcel(inpolygonResults, ChooseFolder());
                MessageBox.Show("Результат сохранен ");
            }
            else
                MessageBox.Show("Сначала выберите файл для импорта данных");
         }

        private void gmap_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                gmap.Overlays.Remove(PositionsForUser);
                gmap.Overlays.Remove(PositionForCicle);

                PositionsForUser.Clear();
                PositionForCicle.Clear();

                gmap.Overlays.Add(PositionsForUser);
                gmap.Overlays.Add(PositionForCicle);

                //постановка центраьлной точки и получение ее координат
                CircleMiddlePosition_X = gmap.FromLocalToLatLng(e.X, e.Y).Lng;
                CircleMiddlePosition_Y = gmap.FromLocalToLatLng(e.X, e.Y).Lat;

                //запоминание центральной координаты для сектора
                CentralCoordinate = new PointLatLng(CircleMiddlePosition_Y, CircleMiddlePosition_X);

                //черная точка в центре круга
                GMarkerGoogle MarkerWithCirclePosition = new GMarkerGoogle(new PointLatLng(CircleMiddlePosition_Y, CircleMiddlePosition_X), GMarkerGoogleType.blue);
                PositionsForUser.Markers.Add(MarkerWithCirclePosition);

                //отрисовка круга после каждого нажатия пкм
                CreateCircle(CircleMiddlePosition_Y, CircleMiddlePosition_X, Convert.ToDouble(numericUpDown1.Value), 1);
            }
        }


        private void button3_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Images|*.png;*.bmp;*.jpg";
            sfd.Title = "Выберите путь для сохранения";
            ImageFormat format = ImageFormat.Png;
            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string ext = System.IO.Path.GetExtension(sfd.FileName);
                switch (ext)
                {
                    case ".jpg":
                        format = ImageFormat.Jpeg;
                        break;
                    case ".bmp":
                        format = ImageFormat.Bmp;
                        break;
                }
                string fname = sfd.FileName;
                try
                {

                    Image b = gmap.ToImage();

                    b.Save(fname, format);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

            }
        }
    }
}
