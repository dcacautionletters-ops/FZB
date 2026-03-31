IndexError: This app has encountered an error. The original error message is redacted to prevent data leaks. Full error details have been recorded in the logs (if you're on Streamlit Cloud, click on 'Manage app' in the lower right of your app).
Traceback:

File "/mount/src/fzb/FZR.py", line 99, in <module>
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
         ~~~~~~~~~~~~~~^^^^^^^^^^^^^^^^^^^^^^^^^^^
File "/home/adminuser/venv/lib/python3.14/site-packages/pandas/io/excel/_base.py", line 1353, in __exit__
    self.close()
    ~~~~~~~~~~^^
File "/home/adminuser/venv/lib/python3.14/site-packages/pandas/io/excel/_base.py", line 1357, in close
    self._save()
    ~~~~~~~~~~^^
File "/home/adminuser/venv/lib/python3.14/site-packages/pandas/io/excel/_openpyxl.py", line 110, in _save
    self.book.save(self._handles.handle)
    ~~~~~~~~~~~~~~^^^^^^^^^^^^^^^^^^^^^^
File "/home/adminuser/venv/lib/python3.14/site-packages/openpyxl/workbook/workbook.py", line 386, in save
    save_workbook(self, filename)
    ~~~~~~~~~~~~~^^^^^^^^^^^^^^^^
File "/home/adminuser/venv/lib/python3.14/site-packages/openpyxl/writer/excel.py", line 294, in save_workbook
    writer.save()
    ~~~~~~~~~~~^^
File "/home/adminuser/venv/lib/python3.14/site-packages/openpyxl/writer/excel.py", line 275, in save
    self.write_data()
    ~~~~~~~~~~~~~~~^^
File "/home/adminuser/venv/lib/python3.14/site-packages/openpyxl/writer/excel.py", line 89, in write_data
    archive.writestr(ARC_WORKBOOK, writer.write())
                                   ~~~~~~~~~~~~^^
File "/home/adminuser/venv/lib/python3.14/site-packages/openpyxl/workbook/_writer.py", line 150, in write
    self.write_views()
    ~~~~~~~~~~~~~~~~^^
File "/home/adminuser/venv/lib/python3.14/site-packages/openpyxl/workbook/_writer.py", line 137, in write_views
    active = get_active_sheet(self.wb)
File "/home/adminuser/venv/lib/python3.14/site-packages/openpyxl/workbook/_writer.py", line 35, in get_active_sheet
    raise IndexError("At least one sheet must be visible")
