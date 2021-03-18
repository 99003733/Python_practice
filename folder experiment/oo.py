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

        
def search_data():
    path_file = {"D:\makefile":[] , "D:\Python Course": [], "D:\MININT" :[]}
    #keys =  list(path.keys())
    for path in path_file.keys():
        print(path)
        #list_files = []
        for f_name in os.listdir(path):
            if f_name.endswith('.xlsx'):
                path_file[path].append(f_name)
        #print(list_files)        
