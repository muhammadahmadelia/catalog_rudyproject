RD /S /Q __pycache__
RD /S /Q build
RD /S /Q dist
del /f safilo.spec
pyinstaller -F -i image_1.ico rudyproject.py
