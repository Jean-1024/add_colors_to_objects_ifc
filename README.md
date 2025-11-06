# add_colors_to_objects_ifc
Adding specific colors to each ifc objects. Get related color code from excel sheet. The code works on ifc file that each object has one body representation and each body representation has one item.

## Add color
### Process
1. Have the 'Bonsai' add-on installed in your Blender
2. Open the ifc file (```Sample_System_Filters.ifc```) in Blender
3. Copy and paste the code from ```ifc_color.py``` and run in 'Scripting' in Blender
   Or just download and open this script in 'Scripting'
4. Make sure you adjust the file path for 'XLSX_PATH' for your excel sheet. (Your file name should be the same as ```SimpleBIM_Type_Filter.xlsx```, just change the path)
5. Run the script
6. Save to a new ifc file and open the new ifc file to check the result.
7. If you use the ```Sample_System_Filters.ifc``` as your input file, you should get result the same as ```output_color.ifc```


## Check this notion link for more information: 
https://cottony-tailor-ab7.notion.site/ifc-projects-299c76371e5980cf82afc5f97faa5713?source=copy_link 
