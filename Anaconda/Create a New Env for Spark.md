# Setting up PySpark, Java, and Hadoop

## Step 1: Install Java

Download and Install Java 17 (not 23).

1. Go to [[https://www.oracle.com/java/technologies/downloads/#java17]]
2. Select your OS
3. Click the file to Download.
    * For macOS-Intel(x64) download the x64 DMG Installer.
    * For Windows download the x64 Installer (ends in EXE).
4. If you haven't already done so, create an Oracle Account (not a Cloud Account). Fill it in, verify your email address, etc.
5. Once logged in, go back to the Java 17 Downloads and click the file again.
6. Install Java 17 on your machine.
7. In Windows, edit your system environment variables and add one called JAVA_HOME, set the value to the path to your Java 17 folder, it is probably: C:\Program Files\Java\jdk-17
8. On the Mac, set JAVA_HOME as follows, in Terminal, enter:

    ```bash
        export JAVA_HOME=$(/usr/libexec/java_home)
    ```

9. On Windows, go to the [GitHub](https://github.com/steveloughran/winutils/)
10. Click into the hadoop-3.0.0 folder.
11. Click into the bin folder.
12. Click into the winutils.exe file.
13. Click the download button in the top right to download the winutils.exe file.
14. Move the winutils.exe file to the "spark-3.5.3-bin-hadoop3\bin" folder.
15. Then set the following environment variables too:

    ```bash
    set HADOOP_HOME = "C:\Program Files\spark-3.5.3-bin-hadoop3"
    set SPARK_HOME = "C:\Program Files\spark-3.5.3-bin-hadoop3"
    set PYTHONPATH = "C:\Program Files\spark-3.5.3-bin-hadoop3\python"
    set PYSPARK_DRIVER_PYTHON = "jupyter"
    set PYSPARK_DRIVER_PYTHON_OPTS = "notebook"
    ```

16. Now you need to install all the other bits. In Windows, go to an Anaconda Prompt. On the Mac, go to Terminal.
17. Then, run these commands, one at a time.

```bash
conda activate base
conda install menuinst
conda create --name spark_env python=3.9 
conda activate spark_env
conda install ipykernel
conda install jupyter
conda install pandas
conda install numpy
conda install scipy
conda install scikit-learn
conda install seaborn
conda install matplotlib
python -m ipykernel install --user --name=spark_env --display-name "Python (spark_env)"
pip install findspark
pip install pyspark==3.5.3
pip install pyarrow
conda install -c conda-forge fastparquet
pip install plotly
```
