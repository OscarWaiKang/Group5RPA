def get_excel_file_paths(text_file):
    with open(text_file, 'r') as file:
        paths = file.readlines()
    return paths[0].strip(), paths[1:]  # Returns the first path and any others