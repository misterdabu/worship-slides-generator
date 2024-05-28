import os
from fuzzywuzzy import fuzz

#get the root directory, file type and file name
root_dir = input("Enter the root directory for your seach: ")
#file_types = input("Enter the file endings to look for (seperate by spaces) (Empty = ALL): ")
fuzzy_search = input("Enter a fuzzy search query (Empty = None): ")

#file_types = file_types.split(" ")
file_types = ['pptx', 'ppt']

for root, dirs, files in os.walk(root_dir):
    for name in files:
        if name.endswith(tuple(ft for ft in file_types)) or file_types[0] == "":
            if fuzz.partial_token_sort_ratio(fuzzy_search.lower(), name.lower()) > 70 or fuzzy_search =="":
                print(root + os.sep + name)