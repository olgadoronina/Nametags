import os
import pandas as pd

#############################################
# Defining paths
#############################################
base_folder = './'      # working directory
# path to .tex template file
tex_template_folder = os.path.join(base_folder, 'template/')
tex_template = os.path.join(tex_template_folder, 'tex_template.tex')
# path to created .tex file with generated nametags (logo file should be in this directory)
tex_output_folder = os.path.join(base_folder, 'tex_output/')
if not os.path.exists(tex_output_folder):
    os.mkdir(tex_output_folder)
tex_output = os.path.join(tex_output_folder, 'tex_output.tex')
logo_file = os.path.join(tex_output_folder, 'logo.png')
if not os.path.exists(logo_file):
    print('Warning: You need to put logo.png file in tex_output/ folder')
# input excel file
excel_file = os.path.join(base_folder, 'Nametag.xlsx')
################################################################
#
################################################################
# read in data
df = pd.read_excel(excel_file, names=['first_name', 'last_name', 'position', 'organization', 'department'], header=None)
# sort rows by last name
df.sort_values(by=['last_name', 'first_name'], inplace=True)
# remove identical rows
df.drop_duplicates(inplace=True)
# replace NaNs with empty string
df.fillna('', inplace=True)
df.department = df.department.str.replace('&', '\&')
df.organization = df.organization.str.replace('&', '\&')
df = df.reset_index(drop=True)

number_nametags = len(df.index)

# read in template file
with open(tex_template, 'r') as f:
    text = f.readlines()
new_text = text[:32]    # save template lines up to (including) \begin{document} to copy in output file

for i in range(number_nametags):

    # if name and surname are too long(longer then 22 characters) split in two lines
    if len(df.last_name[i]) + len(df.first_name[i]) < 22:
        name_line = '\t\confpin{{{} {}}}{{}}'.format(df.first_name[i], df.last_name[i])
    else:
        if len(df.first_name[i]) > 18 or len(df.last_name[i]) > 18:
            print("Warning!: {} {} name is too long for one line".format(df.first_name[i], df.last_name[i]))
        name_line = '\t\confpin{{{}}}{{{}}}'.format(df.first_name[i], df.last_name[i])

    position_line = '{{{}}}'.format(df.position[i])
    if len(df.position[i]) > 34:
        print("Warning!: {} {} position name is too long for one line".format(df.first_name[i], df.last_name[i]))

    department_line = '{{{}}}'.format(df.department[i])
    if len(df.department[i]) > 39:
        print("Warning!: {} {} department name is too long for one line".format(df.first_name[i], df.last_name[i]))

    # if organization name is too long(longer then 35 characters) split in two lines
    words = df.organization[i].split()
    line1, line2 = '', ''
    for w in words:
        if len(line1+' ' + w) < 35:
            line1 += ' ' + w
        else:
            line2 += ' ' + w
    organization_line = '{{{}}}{{{}}}'.format(line1, line2)
    if len(line1) > 35 or len(line2) > 35:
        print("Warning!: {} {} Organization name is too long for one line".format(df.first_name[i], df.last_name[i]))

    # create a string line for ./texfile
    new_text.append(name_line+position_line+department_line+organization_line+'\n')
# add \end{document}
new_text.append(text[-1])
# write ./tex file
with open(tex_output, 'w') as f:
    f.writelines(new_text)
