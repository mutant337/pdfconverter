@echo off

set "home=C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python311"
set "include_system_site_packages=false"
set "version=3.11.4"
set "executable=C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python311\python.exe"
set "command=C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python311\python.exe -m venv C:\Users\%USERNAME%\new_venv"

(
  echo home = %home%
  echo include-system-site-packages = %include_system_site_packages%
  echo version = %version%
  echo executable = %executable%
  echo command = %command%
) > "new_venv\pyvenv.cfg"

echo pyvenv.cfg file created successfully!
