# Hazard-Assessment-CAS-Lookup
This Python project takes a text file with a list of CAS number for the chemicals in your inventory and scrapes Sigma-Aldrich website to retrieve a compilation of Hazard statements (e.g. H319), Precautionary statements (e.g. P120) and recommendations for Personal Protective Equipment. It outputs (4) HTML files, (1) folder with SDS PDF files and (1) spreadsheet as a result.

## Warning
Data may not be exhaustive and/or up-to-date. Do not use this tool as your primary hazard assessement methodoly. Instead, use it to check the results of your existing hazard assessment and identify gaps.

## Install
I recommend that you use a Python environment manager such as Anaconda or Miniconda for this project.  
Anaconda: https://www.continuum.io/downloads  
Miniconda (uses less disk space): http://conda.pydata.org/miniconda.html  
You can then clone my environment into your conda install. That will satisfy the dependencies listed below and get you ready in no time.  
Tutorial at: http://conda.pydata.org/docs/using/envs.html#use-environment-from-file  
Command:
```python
conda env create -f CASlookup.yml
```

## Dependencies (version tested)
python (3.5.1)  
pandas (0.18.1)  
numpy (1.11.0)  
scipy (0.17.1)  
beautifulsoup4 (4.4.1)  
xlsxwriter (0.9.2)  
setuptools (23.0.0)  
python-dateutil (2.5.3)  
pytz (2016.4)  
selenium (2.53.6)  
cssselect (0.9.2)  
mkl (11.3.3)  
six (1.10.0)  
vs2008_runtime (9.00.30729.1)  
vs2015_runtime (14.0.25123)  

## Input
CAS-list.txt : required user editable, (1) CAS number per line, nothing else, empty line ok  
H2P.txt : required, refrain from editing, may break the code  
P-statements.txt : required, refrain from editing, may break the code  
H-statements.txt : required, refrain from editing, may break the code  

## Output
Hazards.html : sorted list of unique H-statements found for the CAS numbers in your list  
Precautions.html : sorted list of unique P-statements found for the CAS numbers in your list  
PPE.html : sorted list of unique PPE recommendations found for the CAS numbers in your list  
Inventory.html : sorted list of chemicals and properties found for the CAS numbers in your list  
Hazard Assessment.xlsx : all data, no formatting  
SDS folder : all SDS found on Sigma-Aldrich website and downloaded as PDF files for the CAS numbers in your list  

## Usage
1 - Customize the CAS list text file  
2 - Activate CASlookup environment in Conda
```python
activate CASlookup  
```
3 - Run
```python
python hazard-assessment-cas-lookup.py
```
4 - Browse the output files  
