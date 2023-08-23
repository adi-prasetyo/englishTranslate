rem Run a Python script in a given conda environment from a batch file.


rem Define here the path to your conda installation
set CONDAPATH=C:\Users\adipr\miniconda3
rem Define here the name of the environment
set ENVNAME=translate

rem The following command activates the environment.
set ENVPATH=%CONDAPATH%\envs\%ENVNAME%

rem Activate the conda environment
rem Using call is required here, see: https://stackoverflow.com/questions/24678144/conda-environments-and-bat-files
call %CONDAPATH%\Scripts\activate.bat %ENVPATH%

set FILENAME=englishIngredientsTranslate

rem Run a python script in that environment
python %FILENAME%.py

rem Deactivate the environment
call conda deactivate