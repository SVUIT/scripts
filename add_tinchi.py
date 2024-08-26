import pandas as pd
import os

# Get list of file paths
def get_all_filenames(root_dir):
    filenames = []
    for root, dirs, files in os.walk(root_dir):
        for file in files:
            filenames.append(os.path.join(root, file))  # Get full path
            # filenames.append(file)  # Get just the filename without the path
    return filenames

df = pd.read_excel('thongketinchi.xlsx', sheet_name='Sheet1')
root_directory = 'docs'  # Replace with your top-level directory
file_list = get_all_filenames(root_directory)

for file in file_list:
    # Get filename
    filename_without_extension = os.path.splitext(os.path.basename(file))[0]
    print(filename_without_extension)
    result = df[df['Môn'] == filename_without_extension.upper()]
    if not result.empty:
        # Get tinchi in file excel    
        tong = result['Tổng'].values[0]
        lithuyet = result['Lý thuyết'].values[0]
        thuchanh = result['Thực hành'].values[0]

        with open(f'{file}', 'r', encoding='utf-8', errors='replace') as reading_file:
            lines = reading_file.readlines()
        # Find the line containing a specific keyword and insert after it
        keyword = "## Mô tả môn học"
        new_text = f"\n### Số tín chỉ: {tong}\n- Lí thuyết: {lithuyet}\n- Thực hành: {thuchanh}\n"
        # Add new text 
        for i, line in enumerate(lines):
            if keyword in line:
                lines.insert(i + 1, new_text)
                break
        # Write the modified content back to the file
        with open(f'{file}', 'w', encoding='utf-8') as writing_file:
            writing_file.writelines(lines)
