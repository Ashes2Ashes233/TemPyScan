Quick Start manual

Step 1: Connect the device

Open the software and you will see the main connection interface.



In the "Device TCP/IP Address" input box, fill in the IP address of your Keithley instrument.



Based on your hardware wiring, select whether the device you are connecting to corresponds to 1-80 or 81-160 in the "Channel Range".



Click the 'Connect' button. If successful, the status will change to green "Connected" and display the device ID. At this point, the "Continue to Test" button will become available.



If you need to disconnect, you can click the "Disconnect" button.



Click the "About" button to access the GitHub homepage of the project, get help or contact developers.



Step 2: Operation and Monitoring

After confirming the device connection, click "Continue to Test" to enter the main running interface.



Click the "Start" button in the upper left corner to start data collection. At this point, the button will become unavailable, and the 'Stop' button will become available.



Data table (left side):



The software will periodically refresh the data of all channels according to the "read interval" you set.



Online editing: You can double-click on the Location or Threshold (° C) cell in any channel row at any time, enter new information, and press Enter to save. This will update the configuration of the channel in real-time.



Curve chart panel (top right):



In the "Channels for Plot/Report" input box, enter the channel number you want to view. Supports single (e.g. 5), multiple (e.g. 5, 8, 12), or range (e.g. 1, 10, 15) inputs.



In the "Time Range (s)" input box, you can specify a time range (relative to the number of seconds at the start of the test), such as Start: 60, End: 300. If left blank, the complete time range will be displayed.



Click the 'Update Plot' button, and the chart on the right will immediately update to the curve of the channel you specified during the specified time period.



Zoom and Pan: Use the Zoom icon in the Matplotlib navigation toolbar below the chart to select zoom in, and use the Pan arrow icon to pan the view.



Step 3: Generate report and export data

After the test is completed (or still in progress), click the "Stop" button to stop data collection.



In the Channels for Plot/Report and Time Range (s) input boxes in the chart area, fill in the channels and time ranges that you want to be included in the report.



In the Report Notes area at the bottom right corner, fill in the "Phenomenon" and "Remarks" for this test.



Click on the 'Proceed to Report Settings' button.



You will be redirected to the report configuration interface. Here, please fill in or modify all report metadata such as Test Name, Tester, Sample number, etc.



After confirming that all information is correct, click the "Confirm\&Generate Report" button.



The program will pop up a file save dialog box. You only need to select the path and specify the file name (no suffix required).



After completion, the program will automatically generate two files in the path you have selected:



A PDF summary report with both graphics and text.



An Excel (. xlsx) spreadsheet containing detailed raw data.



If you want to generate another report with a different title or channel based on the same test data, simply repeat steps 2-8.

