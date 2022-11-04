# Google_Admin_Import_Photos
Formats Synergy Export to GAM Google Admin format.

<H1>Prerequisites</h1>
<ul><li> <a href="https://openpyxl.readthedocs.io/en/stable/">openpyxl</li>
<li><a href="https://docs.python.org/3/library/tkinter.html">tkinter</li>
<li><a href="https://pandas.pydata.org/">pandas</li></ul>



<H1>Directions</h1>

1) Download the "Google Photos" query from Synergy and save to downloads folder
![](https://github.com/aaronzech/images/blob/main/Screenshot_231.png)

2) Run the python script "formatter.py"

3) When the program runs it will prompt the user for the Synergy query

4) Look for the <b>GooglePhotos_GAM.csv</b> file that was created by the program.

5) Open GAM 

6) Run the following command

  ```sh
  gam csv {FILE_LOCATION} gam user ~Email update photo ~File_Location
  ```
