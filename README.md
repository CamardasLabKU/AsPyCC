# AsPyCC
Repository dedicated to the AsPyCC framework for automatization of design, sizing and simulation of Post-Combustion Caron Capture (PCC) units using renewable Ammonia.

## AsPyCC V.1.0

The repository consists of 3 files:

- **AsPyCC.py:** Main code for the AsPyCC framework. This file holds the complete implementation of the AsPyCC framework.
- **Data_generation.ipynb:** This Jupyter notebook holds the procedure to generate the samples for different industries.
- **Flue_gas_db.xlsx:** This Excel files holds a template for the structure of the input for Data_generation.ipynb.

## What is the AsPyCC framework?

The proposed framework for design and simulating PCC units is implemented in Python through a Component Object Model (COM) connection, enabling automated analysis to ensure that column dimensions and process variables meet performance criteria. The Python implementation allows for the integration of process heuristics and specific performance constraints. The absorber is designed to achieve a target CO₂ capture efficiency while maintaining column flooding within acceptable limits and adhering to a maximum allowable diameter, ensuring proper hydraulic performance. The stripper and make-up stream are configured to maintain the same CO₂ loading in the recycle stream as in the lean solvent.

For the framework to work, an Aspen Plus license is required.

## License

According to GitHub documentation:

A repository without a license, the default copyright laws apply, meaning that you retain all rights to your source code and no one may reproduce, distribute, or create derivative works from your work. 

The framework can be used for educational and research purposes. Any commercial use of the tool is prohibited.

## Citation

Pending.
