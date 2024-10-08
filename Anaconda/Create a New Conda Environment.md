# New Conda Environment Setup

## To list the environments you have

``` Bash
conda info -e
```

## Initial Setup for a new Data environment

``` Bash
conda create --name DataViz python=3.12
conda activate DataViz
conda install pandas
conda install console_shortcut
conda install -c conda-forge jupyterlab
conda install -c conda-forge matplotlib
pip install python-dotenv
pip install jupyter --upgrade
pip install ipympl
conda install scipy
conda install scikit-learn
conda install -c pyviz hvplot geoviews
pip install census
pip install citipy
conda install -c anaconda sqlalchemy
conda install -c anaconda sqlite
pip install flask
pip install "splinter[selenium4]"
pip install bs4
pip install html5lib
pip install lxml
pip install pymongo
```

## Create the kernel for VS Code/Jupyter

``` Bash
python -m ipykernel install --user --name <other-env> --display-name "Python <ver.> (<other-env>)"
For example,
python -m ipykernel install --user --name DataViz --display-name "Python 3.12 (DataViz)"
```
### To remove a Kernel from Jupyter, simply run the following code:

``` Bash
jupyter kernelspec uninstall "Python 3.7.1 64-bit"
```

### To list the Kernel's you have defined

``` Bash
jupyter kernelspec list
```

## Occasionally:

Run this from time-to-time to update Anaconda

``` Bash
conda update -n base -c defaults conda
```

Run this from time-to-time to update pip

``` Bash
python.exe -m pip install --upgrade pip
```

## To Remove a Corrupted Environment

``` Bash
conda env remove -n ENV_NAME
```

### To Launch Jupyter Lab

PC:  From the Anaconda Prompt (DataViz):

``` Bash
jupyter lab
```

Mac:  From the Terminal:

``` Bash
conda activate DataViz
jupyter lab
```
