# generator_script

A notice or certificate generator script in python.\
Author: Lien, Zi-Li

## Starting the Program

### Run with Python

First, install all dependencies which are listed on `requirements.txt` by running the following line in terminal:

```
pip install -r requirements.txt
```

Next, run the python file.

### Run with Executable File

Download `generator.exe`: [Download Here](https://drive.google.com/file/d/1pyy0mBgzNAFMTbaUMiPG6PApDa8YgOhq/view?usp=sharing)\
After download, simply click on `generator.exe`. It can take a while to startup, and note that it can **ONLY** be executed on Windows system.

## Using the Program

Generally, follow the steps listed on the UI.

1. Select a template file(`.docx`, document file only) with placeholder texts(texts to be replaced) in it.
2. Select a data file(`.xlsx`, Excel file only) containing title and data.
3. Enter the number of placeholder texts, the placeholder texts, and the corresponding title. Note that for multiple placeholder texts and titles, entries should be separated with **ONLY comma**(`,`), not space nor other symbols. Then, save it.
4. Enter the desired file name to save with and the title for different files to be named with. The generated files will then be named as `[file_name]_[data.title[i]].docx/.pdf`. Note that the title should be identical to one of that entered in step 3.\
   Next, select a file where the generated files are saved. Then, save the data entered above.
5. Select the file type to be exported, either `.docx` or `.pdf`.
6. Click on the `Start execute` button and wait for the files to be generated. Note that `.pdf` files can take way longer time to generate than `.docx` files.

For more detailed example, refer to `manual.pdf`.
