# Installation

```bash
pip install pywin32
pip install xlwinds
pip install flask
python run_app.py
```

# Requirements
Only works on Windows machine and Excel must be installed.  
Change path in _api/config.py_ file to point to the _import_ folder.

# Donation Button
[![Donate](https://img.shields.io/badge/Donate-PayPal-green.svg)](https://www.paypal.com/cgi-bin/webscr?cmd=_xclick&business=kyeoses%40gmail%2ecom&lc=SG&item_name=Donate&amount=0%2e00&currency_code=USD&button_subtype=services&bn=PP%2dBuyNowBF%3abtn_buynowCC_LG%2egif%3aNonHosted)


## Notes when running on Server
Server must have Excel installed.  
When running program on server, it will still work while running _run_app.py_ locally.  
But when hosted on IIS, it might encounter errors of such :
```python
Traceback (most recent call last):
  File "C:\excel_to_pdf\api\Transform_to_PDF.py", line 56, in transform_to_pdf
    wb = excel.Workbooks.Open(file_path)
  File "<COMObject <unknown>>", line 8, in Open
pywintypes.com_error: (-2147352567, 'Exception occurred.', (0, 'Microsoft Excel', "Sorry, we couldn't find C:\\excel_to_pdf\\api\\import\\test.xlsx. Is it possible it was moved, renamed or deleted?", 'xlmain11.chm', 0, -2146827284), None)
```

To workaround this issue, create a _Desktop_ folder under the relevant folder:
- *Windows Server x64* - C:\Windows\SysWOW64\config\systemprofile\Desktop
- *Windows Server x86* - C:\Windows\System32\config\systemprofile\Desktop  

Refer to this [link](https://stackoverflow.com/questions/12571985/task-scheduler-wont-run-python-script-pywin32-to-open-excel-how-to-get-more) for more info.

Take note that even if the server is 64bit, python may be running on 32bit, so the folder will have to be on the _System32_ path.


# Execution
Run *POST* on this endpoint to test the API, it should return a pdf version of the _test.xlsx_ in this repository.
```python
https://<server>/api/v1/excel_to_pdf?test
```

For real execution, run *POST* on this endpoint, and provide url params for _sheet_ to specify the Sheet name that needs conversion to pdf.
```python
https://<server>/api/v1/excel_to_pdf?sheet=Sheet1
```

For commandline runs, use _cpdfy.py_ and provide these arguments:
- f : absolute file input path
- s : sheet name to convert
- d : absolute file destination path
```bash
python cpdfy.py -f "C:/excel_to_pdf/api/import/test.xlsx" -s "Sheet1" -d "C:/excel_to_pdf/api/import/test.pdf"
```
