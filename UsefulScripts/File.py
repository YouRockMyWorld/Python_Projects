import glob


def get_all_files(dir_path, ext = '*.*'):
    '''
    得到指定目录下的所有文件列表
    :param dir_path: 文件夹路径，如C:\\Users\\Administrator\\Desktop\\test
    :param ext: 过滤文件名, 如'*.txt'
    :return: 文件完整路径列表
    '''
    if dir_path.endswith('\\'):
        filter_ext = dir_path +ext
    else:
        filter_ext = dir_path + '\\' +ext
    filepaths = []
    for filename in glob.glob(filter_ext):
        filepaths.append(filename)
    return filepaths