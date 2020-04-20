# Mathematical Protocol for Opioid Weaning (MPOW)
This repository holds the data and the Python data-analysis utilities for a collaborative
research effort conducted in the Physical Medicine and Rehabilitation group, contributors include: 
Dr. Amir Ahmadian, Dr. William Lian, Dr. Jaclyn Nguyen, Dr. Shannon B. Juengst, Dr. Ugo Bitussi, 
James Kennington, Dr. Fatma Gul, and Dr. Kathleen R Bell. Publication currently under review.

## Data Analysis

### Raw Data
The data were originally contained in an excel spreadsheet in a pivoted format, where each row
represented a day and the columns represented successive intraday measurements of pain scores or 
opioid consumption. This raw format is still available in `data/MPOW-Data-20190206.xlsx`.

### Normalized Data
The data were extracted from the spreadsheet and normalized into relational form, in which 
each row represents a single measurement (either daily or intraday depending on the quantity
being measured). The normalization and extraction utilities are available in `mpow/load_data.py`,
and the normalized data have been saved to `data/data.h5` in HDF5 format.

### Regression
Though the regression helper functions are in `mpow/regression.py`, the bulk of the analysis
is in the Jupyter notebook `analysis.ipynb`. This notebook can be accessed by 
running `jupyter notebook` in a properly-setup python console, or by clicking the binder link
[here](https://mybinder.org/v2/gh/JWKennington/collab-mpow/master?filepath=analysis.ipynb). 
