import pypyodbc
import xlsxwriter


# 建立数据库连接
def mdb_conn(db_name, password=""):
    """
    功能：创建数据库连接
    :param db_name: 数据库名称
    :param db_name: 数据库密码，默认为空
    :return: 返回数据库连接
    """
    str = 'Driver={Microsoft Access Driver (*.mdb, *.accdb)};PWD' + password + ';DBQ=' + db_name
    conn = pypyodbc.win_connect_mdb(str)
    return conn


# 查询记录
def mdb_sel_all(cur, sql):
    """
    功能：向数据库查询数据
    :param cur: 游标
    :param sql: sql语句
    :return: 查询结果集
    """
    try:
        cur.execute(sql)
        return cur.fetchall()
    except:
        return []


def mdb_sel_one(cur, sql):
    """
    功能：向数据库查询数据
    :param cur: 游标
    :param sql: sql语句
    :return: 查询结果
    """
    try:
        cur.execute(sql)
        return cur.fetchone()
    except:
        return None


def ExtractData(table, mdb_connection):
    '''
    提取数据
    :param table:
    :param mdb_connection:
    :return: 数据列表
    '''
    result_data = []
    line_table = table + 'L'
    point_table = table + 'P'
    cur = mdb_connection.cursor()
    pipeline_sql = "select * from %s;" % line_table
    pipeline_data = mdb_sel_all(cur, pipeline_sql)
    for item in pipeline_data:
        pipeline = []
        pipeline.append(item[0])  # 起点号
        pipeline.append(item[1])  # 终点号
        pipeline_p_sql = "select 管线起点号, 特征, 附属物, X坐标, Y坐标, 地面高程 from %s where 管线起点号 = '%s';" % (point_table, item[0])
        pipeline_PointData = mdb_sel_one(cur, pipeline_p_sql)
        if pipeline_PointData is None:
            continue
        pipeline.append(pipeline_PointData[1])  # 特征
        pipeline.append(pipeline_PointData[2])  # 附属物
        pipeline.append(pipeline_PointData[3])  # X
        pipeline.append(pipeline_PointData[4])  # Y
        pipeline.append(pipeline_PointData[5] - item[2])  # Z
        pipeline.append(item[6])  # 材质
        pipeline.append(item[8])  # 尺寸

        # 其它属性
        pipeline.append(item[7])  # 埋设方式
        pipeline.append(None if item[9] is None else item[9].strftime('%Y-%m-%d'))  # 建设年代
        pipeline.append(item[10])  # 权属单位
        pipeline.append(item[11])  # 电缆条数
        pipeline.append(item[12])  # 电压值
        pipeline.append(item[13])  # 压力类型
        pipeline.append(item[14])  # 总孔数
        pipeline.append(item[15])  # 已用孔数
        pipeline.append(item[16])  # 套管尺寸材质
        pipeline.append(item[17])  # 道路名称
        pipeline.append(item[18])  # 排水流向
        pipeline.append(item[19])  # 备注
        pipeline.append(item[0] + '-' + item[1])  # 管道名称

        # print(pipeline)
        result_data.append(pipeline)

        pipeline_reverse = []
        pipeline_reverse_p_sql = "select 管线起点号, 特征, 附属物, X坐标, Y坐标, 地面高程 from %s where 管线起点号 = '%s';" % (
            point_table, item[1])
        pipeline_reverse_PointData = mdb_sel_one(cur, pipeline_reverse_p_sql)
        if pipeline_reverse_PointData is None:
            continue
        pipeline_reverse.append(item[1])
        pipeline_reverse.append(item[0])
        pipeline_reverse.append(pipeline_reverse_PointData[1])  # 特征
        pipeline_reverse.append(pipeline_reverse_PointData[2])  # 附属物
        pipeline_reverse.append(pipeline_reverse_PointData[3])  # X
        pipeline_reverse.append(pipeline_reverse_PointData[4])  # Y
        pipeline_reverse.append(pipeline_reverse_PointData[5] - item[3])  # Z
        pipeline_reverse.append(item[6])  # 材质
        pipeline_reverse.append(item[8])  # 尺寸

        # 其它属性
        pipeline_reverse.append(item[7])  # 埋设方式
        pipeline_reverse.append(None if item[9] is None else item[9].strftime('%Y-%m-%d'))  # 建设年代
        pipeline_reverse.append(item[10])  # 权属单位
        pipeline_reverse.append(item[11])  # 电缆条数
        pipeline_reverse.append(item[12])  # 电压值
        pipeline_reverse.append(item[13])  # 压力类型
        pipeline_reverse.append(item[14])  # 总孔数
        pipeline_reverse.append(item[15])  # 已用孔数
        pipeline_reverse.append(item[16])  # 套管尺寸材质
        pipeline_reverse.append(item[17])  # 道路名称
        pipeline_reverse.append(item[18])  # 排水流向
        pipeline_reverse.append(item[19])  # 备注
        pipeline_reverse.append(item[1] + '-' + item[0]) #管道名称

        # print(pipeline_reverse)
        result_data.append(pipeline_reverse)

    cur.close()

    return result_data


def ExtractData_appendage(table, mdb_connection):
    '''
    修改管线点高程，增加埋深，适用于附属物数据
    :param table:
    :param mdb_connection:
    :return:
    '''
    result_data = []
    line_table = table + 'L'
    point_table = table + 'P'
    cur = mdb_connection.cursor()
    pipeline_sql = "select * from %s;" % line_table
    pipeline_data = mdb_sel_all(cur, pipeline_sql)
    for item in pipeline_data:
        pipeline = []
        pipeline.append(item[0])  # 起点号
        pipeline.append(item[1])  # 终点号
        pipeline_p_sql = "select 管线起点号, 特征, 附属物, X坐标, Y坐标, 地面高程 from %s where 管线起点号 = '%s';" % (point_table, item[0])
        pipeline_PointData = mdb_sel_one(cur, pipeline_p_sql)
        if pipeline_PointData is None:
            continue
        pipeline.append(pipeline_PointData[1])  # 特征
        pipeline.append(pipeline_PointData[2])  # 附属物
        pipeline.append(pipeline_PointData[3])  # X
        pipeline.append(pipeline_PointData[4])  # Y
        pipeline.append(pipeline_PointData[5])  # Z
        pipeline.append(item[6])  # 材质
        pipeline.append(item[8])  # 尺寸

        # 其它属性
        pipeline.append(item[7])  # 埋设方式
        pipeline.append(None if item[9] is None else item[9].strftime('%Y-%m-%d'))  # 建设年代
        pipeline.append(item[10])  # 权属单位
        pipeline.append(item[11])  # 电缆条数
        pipeline.append(item[12])  # 电压值
        pipeline.append(item[13])  # 压力类型
        pipeline.append(item[14])  # 总孔数
        pipeline.append(item[15])  # 已用孔数
        pipeline.append(item[16])  # 套管尺寸材质
        pipeline.append(item[17])  # 道路名称
        pipeline.append(item[18])  # 排水流向
        pipeline.append(item[19])  # 备注
        pipeline.append((item[2] + 0.5)*1000)  # 附属物埋深
        # print(pipeline)
        result_data.append(pipeline)

        pipeline_reverse = []
        pipeline_reverse_p_sql = "select 管线起点号, 特征, 附属物, X坐标, Y坐标, 地面高程 from %s where 管线起点号 = '%s';" % (
            point_table, item[1])
        pipeline_reverse_PointData = mdb_sel_one(cur, pipeline_reverse_p_sql)
        if pipeline_reverse_PointData is None:
            continue
        pipeline_reverse.append(item[1])
        pipeline_reverse.append(item[0])
        pipeline_reverse.append(pipeline_reverse_PointData[1])  # 特征
        pipeline_reverse.append(pipeline_reverse_PointData[2])  # 附属物
        pipeline_reverse.append(pipeline_reverse_PointData[3])  # X
        pipeline_reverse.append(pipeline_reverse_PointData[4])  # Y
        pipeline_reverse.append(pipeline_reverse_PointData[5])  # Z
        pipeline_reverse.append(item[6])  # 材质
        pipeline_reverse.append(item[8])  # 尺寸

        # 其它属性
        pipeline_reverse.append(item[7])  # 埋设方式
        pipeline_reverse.append(None if item[9] is None else item[9].strftime('%Y-%m-%d'))  # 建设年代
        pipeline_reverse.append(item[10])  # 权属单位
        pipeline_reverse.append(item[11])  # 电缆条数
        pipeline_reverse.append(item[12])  # 电压值
        pipeline_reverse.append(item[13])  # 压力类型
        pipeline_reverse.append(item[14])  # 总孔数
        pipeline_reverse.append(item[15])  # 已用孔数
        pipeline_reverse.append(item[16])  # 套管尺寸材质
        pipeline_reverse.append(item[17])  # 道路名称
        pipeline_reverse.append(item[18])  # 排水流向
        pipeline_reverse.append(item[19])  # 备注
        pipeline_reverse.append((item[3] + 0.5)*1000)  # 附属物埋深
        # print(pipeline_reverse)
        result_data.append(pipeline_reverse)

    cur.close()

    return result_data


def write_data_to_excel(file_path, table, data, data_appendage):
    workbook = xlsxwriter.Workbook(file_path)
    header_format = workbook.add_format(
        {'bold': True, 'font_size': 12, 'font_name': '微软雅黑', 'align': 'center', 'valign': 'vcenter'})
    data_format = workbook.add_format({'font_size': 11, 'font_name': '微软雅黑', 'align': 'center', 'valign': 'vcenter'})

    # 数据表
    worksheet = workbook.add_worksheet(table)
    header = ['管线点号', '连接点号', '管线点特征', '管线点附属物', 'X（m）', 'Y（m）', 'Z（m）', '材质', '尺寸（mm）',
              '埋设方式', '建设年代', '权属单位', '电缆条数', '电压值', '压力类型', '总孔数', '已用孔数', '套管尺寸材质', '道路名称', '排水流向', '备注',
              '管道名称']
    worksheet.write_row(0, 0, header, cell_format=header_format)
    i = 1
    for item in data:
        worksheet.write_row(i, 0, item, cell_format=data_format)
        i += 1

    # 附属物表
    worksheet = workbook.add_worksheet(table + '_appendage')
    header_appendage = ['管线点号', '连接点号', '管线点特征', '管线点附属物', 'X（m）', 'Y（m）', 'Z（m）', '材质', '尺寸（mm）',
                        '埋设方式', '建设年代', '权属单位', '电缆条数', '电压值', '压力类型', '总孔数', '已用孔数', '套管尺寸材质', '道路名称', '排水流向', '备注',
                        '埋深']
    worksheet.write_row(0, 0, header_appendage, cell_format=header_format)
    i = 1
    for item in data_appendage:
        worksheet.write_row(i, 0, item, cell_format=data_format)
        i += 1

    workbook.close()


if __name__ == '__main__':
    table = 'GY'
    MDB_filePath = r'C:\Users\******数据库.mdb'
    excel_path = r'C:\Users\******\%s.xlsx' % table
    conn = mdb_conn(MDB_filePath)

    data = ExtractData(table, conn)
    data_appendage = ExtractData_appendage(table, conn)
    print(len(data))
    print(len(data_appendage))

    conn.close()

    write_data_to_excel(excel_path, table, data, data_appendage)
