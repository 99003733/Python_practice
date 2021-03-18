def search_data():
    path = {"D:\makefile":[] , "D:\Python Course": [], "D:\MININT" :[]}
    keys =  list(path.keys())
    for i in range(len(keys)):
        print(keys[i])
        list_files = []
        for f_name in os.listdir(keys[i]):
            if f_name.endswith('.xlsx'):
                list_files.append(f_name)
        print(list_files)
