# ir_ocv_recorder
Battery IR and OCV Data Logger
A tool designed for battery characterization, providing accurate measurements and automated logging of a battery's Internal Resistance (IR) and Open-Circuit Voltage (OCV). This project is ideal for engineers, researchers, and hobbyists looking to analyze battery health and performance over time.

Features
Accurate Measurements: Precisely calculates Internal Resistance and measures Open-Circuit Voltage.

Automated Data Logging: Automatically records timestamped IR and OCV values to a CSV file for easy analysis.

Configurable Parameters: Allows users to set the logging interval and measurement parameters.

Real-time Monitoring: A simple command-line interface displays live readings as they are captured.

Data Analysis Scripts (Optional): Includes Python scripts to plot and visualize the logged data to track battery degradation.

Tech Stack
Hardware: Designed to run with an Arduino or ESP32 for data acquisition, connected to a simple voltage measurement circuit.

Software:

Python: The main application for controlling the microcontroller and logging data.

Pyserial: For serial communication between the computer and the microcontroller.

Pandas & Matplotlib: For data handling and visualization in the analysis scripts.

How to Run
Set up the hardware: Connect the battery to the measurement circuit as shown in the schematics.

Clone the repository:

git clone [https://github.com/your-username/your-repository-name.git](https://github.com/your-username/your-repository-name.git)
cd your-repository-name

Install dependencies:

pip install -r requirements.txt

Run the logger:

python record_data.py
