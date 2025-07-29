# TemPyScan - Thermocouple Scanning and Reporting Software
Function Description
TemPyScan is a desktop application designed specifically for Keithley 2701 data acquisition/multi-channel comprehensive testing system, aiming to provide a complete solution from multi-channel temperature data acquisition, real-time monitoring to professional report generation. The software enables users to easily complete complex temperature rise tests and automate tedious data organization and report production work through a simple graphical interface.

The core functions include:

Flexible device connectivity: Supports connecting Keithley 2700/2701 hosts via TCPIP (Socket) and can adapt to different hardware configurations by selecting channel ranges (1-80/81-160).

Real time data display and monitoring: Display the current temperature and historical highest temperature of up to 160 channels in real-time in a tabular format, and support independent setting of Location and Threshold for each channel. When the temperature exceeds the threshold, the corresponding data will be highlighted in red as a warning.

Powerful data visualization: Provides an interactive curve chart panel. Users can input the channel number they are interested in at any time (supporting range selection, such as (1,10), 15), and select a specific time range. The software will immediately draw the historical temperature curve of the corresponding channel during the specified time period. The chart comes with a navigation toolbar that supports pan and zoom, making it easy to analyze data details in depth.

Highly flexible reporting system: The original "test configuration report" workflow allows users to return to the parameter configuration interface multiple times after a data collection is completed, modify the metadata of the report (such as test name, sample number, etc.), and generate multiple professional PDF reports with different headers based on the same raw data.

Intelligent grouping drawing: When generating reports, the software automatically groups the user specified channel list into groups of every 8, and generates independent curve graphs with clear titles for each group, perfectly balancing the needs of data overview and detail analysis.

One click data export: At the same time as generating a PDF report, all channels included in the report and complete raw data within a specified time range will be automatically exported as a clear format Excel (. xlsx) file with the same name for offline deeper data analysis.
