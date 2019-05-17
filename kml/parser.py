import xml.etree.cElementTree as ET
import glob, os
import xml.dom.minidom as MN
from CONF import *



def get_all_files(dir_path, ext = '*.*'):
    '''
    得到指定目录下的所有文件列表
    :param dir_path: 文件夹路径，如C:\\Users\\Administrator\\Desktop\\test
    :param ext: 过滤文件名, 如'*.txt'
    :return: 文件完整路径列表
    '''
    filter_ext = os.path.join(dir_path, ext)
    filepaths = []
    for filename in glob.glob(filter_ext):
        filepaths.append(filename)
    return filepaths


def test():
    tree = ET.parse(r'C:\Users\Administrator\Desktop\KML\*****.kml')
    placemark = tree.findall('Document/Placemark')
    print(len(placemark))
    for item in placemark:
        name = item.find('name').text.strip()
        coor = item.find('Point/coordinates').text.strip()
        path = r'C:\Users\Administrator\Desktop\***\%s.txt' % (name)
        with open(path, 'w') as f:
            s = coor.replace(' ', '\n')
            f.write(s)


def get_kml_coordinates(path,xpath_filter):
    with open(path, 'r') as f:
        data = f.read()
        data = data.replace(' xmlns="http://earth.google.com/kml/2.1"','')
    tree = ET.fromstring(data)
    coordinates = tree.find(xpath_filter)
    return coordinates.text.strip()


def create_kml(dir_path):
    root = ET.Element('kml')
    Document = ET.SubElement(root, 'Document')
    name = ET.SubElement(Document, 'name')
    name.text = '******地图数据'
    open_ = ET.SubElement(Document, 'open')
    open_.text = '1'

    #单位工程范围****************************************************************************************
    print('开始输出单位工程范围信息...')
    Folder_danweigongchengfanwei = ET.SubElement(Document, 'Folder')
    name = ET.SubElement(Folder_danweigongchengfanwei, 'name')
    name.text = '单位工程范围'
    files_danweigongcheng = get_all_files(os.path.join(dir_path,'单位工程范围'), '*.kml')
    for f in files_danweigongcheng:
        station_name = os.path.splitext(os.path.basename(f))[0]
        data = get_kml_coordinates(f,'Document/Folder/Placemark/Polygon/outerBoundaryIs/LinearRing/coordinates')

        Placemark = ET.SubElement(Folder_danweigongchengfanwei, 'Placemark')
        name = ET.SubElement(Placemark, 'name')
        name.text = station_name + '范围'
        Style = ET.SubElement(Placemark, 'Style')
        LineStyle = ET.SubElement(Style, 'LineStyle')
        color = ET.SubElement(LineStyle, 'color')
        color.text = 'ff555555'
        width = ET.SubElement(LineStyle, 'width')
        width.text = '1'
        LineString = ET.SubElement(Placemark, 'LineString')
        coordinates = ET.SubElement(LineString, 'coordinates')
        coordinates.text = '\n\t\t\t\t\t' + data.replace('\n','\n\t\t\t\t\t') + '\n\t\t\t\t\t'
        OvStyle = ET.SubElement(Placemark, 'OvStyle')
        TrackStyle = ET.SubElement(OvStyle, 'TrackStyle')
        type = ET.SubElement(TrackStyle, 'type')
        type.text = '5'
        width = ET.SubElement(TrackStyle, 'width')
        width.text = '1'

    #线路****************************************************************************************
    print('开始输出线路信息...')
    Folder_xianlu = ET.SubElement(Document, 'Folder')
    name = ET.SubElement(Folder_xianlu, 'name')
    name.text = '****线路'
    open_ = ET.SubElement(Folder_xianlu, 'open')
    open_.text = '1'

    files_xianlu_shangxing = get_all_files(os.path.join(dir_path,'****线路','上行'), '*.kml')

    Placemark = ET.SubElement(Folder_xianlu, 'Placemark')
    name = ET.SubElement(Placemark, 'name')
    name.text = '上行'
    Style = ET.SubElement(Placemark, 'Style')
    LineStyle = ET.SubElement(Style, 'LineStyle')
    color = ET.SubElement(LineStyle, 'color')
    color.text = 'ffd18802'
    width = ET.SubElement(LineStyle, 'width')
    width.text = '1'
    MultiGeometry = ET.SubElement(Placemark, 'MultiGeometry')

    for f in files_xianlu_shangxing:
        #line_name = os.path.splitext(os.path.basename(f))[0]
        data = get_kml_coordinates(f, 'Document/Folder/Placemark/LineString/coordinates')
        LineString = ET.SubElement(MultiGeometry, 'LineString')
        coordinates = ET.SubElement(LineString, 'coordinates')
        coordinates.text = '\n\t\t\t\t\t\t' + data.replace('\n','\n\t\t\t\t\t\t') + '\n\t\t\t\t\t\t'
    OvStyle = ET.SubElement(Placemark, 'OvStyle')
    TrackStyle = ET.SubElement(OvStyle, 'TrackStyle')
    type = ET.SubElement(TrackStyle, 'type')
    type.text = '5'
    width = ET.SubElement(TrackStyle, 'width')
    width.text = '1'


    files_xianlu_xiaxing = get_all_files(os.path.join(dir_path, '****线路', '下行'), '*.kml')
    Placemark = ET.SubElement(Folder_xianlu, 'Placemark')
    name = ET.SubElement(Placemark, 'name')
    name.text = '下行'
    Style = ET.SubElement(Placemark, 'Style')
    LineStyle = ET.SubElement(Style, 'LineStyle')
    color = ET.SubElement(LineStyle, 'color')
    color.text = 'ffd18802'
    width = ET.SubElement(LineStyle, 'width')
    width.text = '1'
    MultiGeometry = ET.SubElement(Placemark, 'MultiGeometry')

    for f in files_xianlu_xiaxing:
        #line_name = os.path.splitext(os.path.basename(f))[0]
        data = get_kml_coordinates(f, 'Document/Folder/Placemark/LineString/coordinates')
        LineString = ET.SubElement(MultiGeometry, 'LineString')
        coordinates = ET.SubElement(LineString, 'coordinates')
        coordinates.text = '\n\t\t\t\t\t\t' + data.replace('\n', '\n\t\t\t\t\t\t') + '\n\t\t\t\t\t\t'
    OvStyle = ET.SubElement(Placemark, 'OvStyle')
    TrackStyle = ET.SubElement(OvStyle, 'TrackStyle')
    type = ET.SubElement(TrackStyle, 'type')
    type.text = '5'
    width = ET.SubElement(TrackStyle, 'width')
    width.text = '1'


    #单位工程位置****************************************************************************************
    print('开始输出单位工程位置信息...')
    Folder_danweigongcheng = ET.SubElement(Document, 'Folder')
    name = ET.SubElement(Folder_danweigongcheng, 'name')
    name.text = '单位工程'
    file = os.path.join(dir_path, '单位工程','**经纬度.kml')
    data = get_kml_coordinates(file, 'Document/Folder/Placemark/LineString/coordinates')

    datalist = data.split('\n')

    if(len(datalist) == len(*******LIST)):
        print('**数量：'+ str(len(datalist)))
        for i in range(len(datalist)):
            Placemark = ET.SubElement(Folder_danweigongcheng, 'Placemark')
            name = ET.SubElement(Placemark, 'name')
            name.text = *******LIST[i]
            Style = ET.SubElement(Placemark, 'Style')
            IconStyle = ET.SubElement(Style, 'IconStyle')
            color = ET.SubElement(IconStyle, 'color')
            color.text = 'ffffffff'
            scale = ET.SubElement(IconStyle, 'scale')
            scale.text = '1'
            Icon = ET.SubElement(IconStyle, 'Icon')
            href = ET.SubElement(Icon, 'href')
            href.text = 'http://maps.google.com/mapfiles/kml/shapes/bus.png'
            Point = ET.SubElement(Placemark, 'Point')
            coordinates = ET.SubElement(Point, 'coordinates')
            coordinates.text = datalist[i]


    rawText = ET.tostring(root)
    dom = MN.parseString(rawText)
    with open(os.path.join(dir_path,'testout.kml'), 'w', encoding='utf-8') as f:
        dom.writexml(f, '', '\t', '\n', encoding='utf-8')



if __name__ == '__main__':
    create_kml(r'C:\Users\Administrator\Desktop\****\原始数据')
