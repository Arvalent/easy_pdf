# Easy PDF

<p align="center">
    <img src="logo.png" alt="Easy PDF Image" />
</p>

## Welcome!

This repo contains the code of the Easy PDF app, which aims at providing a free and simple way to manipulate PDF files.

At the moment, it is available in French, English and German and it supports two types of operations: 

- merge multiple PDFs into one

- extract some pages from a PDF

To get more details about the app usage and download an installation file (only Windows right now), go [here](https://github.com/Arvalent/easy_pdf_exe). <br><br>


## The code

If you want to contribute or just use the code for your own applications and customize it, here are a few guidelines:

- _Easy PDF_ relies on a few python libraries. Requirements are in the _requirements.txt_ file.

    To install the libraries, it is better to first create a virtual environment. With conda for example, just run:

    ```bash
    conda create -n env_name

    conda activate env_name
    ```

    And to get out of the environment run:
    ```bash
    conda deactivate
    ```

    Once you are in the environment, run:

    ```bash
    pip install -r requirements.txt
    ```

- To launch the app, use the main script _easy_pdf.py_ and run without argument:

    ```bash
    python easy_pdf.py
    ```

- If you want to add some features (buttons, labels, ...), to be consistent with the language handling, you can add new keys in the dictionary that is created in _texts.py_. To retrieve the text inside the __Interface__ class, you can just call:

    ```python
    texts[self.params['language']]['corresponding_key_in_the_texts_dict']
    ```

    If you want to add an entire new language to the app, you are more than welcome!

- If you want to create an executable folder or file, you can use [PyInstaller](https://pyinstaller.org/en/stable/). You can install it simply with:

    ```bash
    pip install -U pyinstaller
    ```
    
    Then, to create the executable, I recommend using:
    
    ```bash
    pyinstaller -w easy_pdf.py
    ```
    
    This way you avoid having a terminal window that opens wih the app. You can remove the -w option though for debug mode.
    
    A bunch of files will be created and you need to copy paste _logo.ico_ and _up_arrow.png_ into ```dist/easy_pdf/```. The _easy_pdf.exe_ will be in the same directory. You can also create a single file instead of a whole directory with:
    
    ```bash
    pyinstaller -w --onefile easy_pdf.py
    ```
    
    However, _PyInstaller_ warns that the app can be slower at the start with this option.
    
    Finally, note that _PyInstaller_ generates executables that are OS specific. For instance if you compile it on Windows, you will get a Windows executable.
    
- For Windows users who want to create an installation file, [Inno Setup](https://jrsoftware.org/isinfo.php) works great. It should be used with the _easy_pdf_ folder that _PyInstaller_ generates.

- Be careful that when you run the app for the first time, a _params.txt_ file is generated in the same folder as either _easy_pdf.py_ or _easy_pdf.exe_ (depending on your usage). This file records the user's preferences in terms of language and default directory. So __it should be deleted if you plan to share the folder with anyone or before using _Inno Setup___. <br><br>


## Planned improvements

- [ ] Create a new button to remove a single file from the list in the merge configuration.

- [ ] Create a PDF viewer for the merge option so that the user can view the PDF by double clicking on the filename or by clicking on Enter.

- [ ] Handle images so that images can be saved as PDF or even merged together in a PDF or even with other PDFs.

- [ ] Create a more flexible third option where you can merge and extract at the same time.

- [ ] Add another language ?
