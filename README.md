# PCB Thermal Analyzer

PCB Thermal Analyzer is a PyAEDT based web-app to simulate PCB thermal response. 
The app uses Ansys Icepak to read in a ECAD file and IDF file and assign boundary conditions 
based on a modified boundary conditions file extracted from the IDF file. 

Download the files on to a desired directory on your computer. Navigate to the directory using a 
terminal or a command prompt and type the following the command to install required packages to 
run the app:

```python
pip install -r requirements.txt
```

Once the required packages have been installed, in the command prompt, point to the directory containing the app files and type in the following command:

```python
python pcb_thermal_analyzer.py
```

