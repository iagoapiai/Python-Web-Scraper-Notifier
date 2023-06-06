## WebScraperNotifier
The Python script performs web scraping on the company's website, accessing each client individually to check if the asset is online. If the parameter changes from offline to online or vice versa, it records the change in a .csv file, which essentially serves as an improvised database, and notifies the parameter change through a message on Discord.

## :warning: Prerequisites

- [Time](https://docs.python.org/3/library/time.html)

- [Pathlib](https://docs.python.org/3/library/pathlib.html)

- [Openpyxl](https://openpyxl.readthedocs.io/en/stable/)

- [Selenium](https://www.selenium.dev/documentation/)

## WHAT THE SCRIPT DO:
- 1° Does the scraping inside the website of each company;
- 2° Compare the information with an improvised database;
- 3° If the parameter changes, it notifies on discord;
- 4° If the current date is the same as the name of the file "horario" it shows online;
- 5° It's in an infinite loop

## OBSERVATION
This code will not work because I removed most of the variables because this data is confidential from the company where I currently work!

## Notification Image
<p align="center">
    <img src="https://github.com/iagoapiai/WebScraperNotifier/assets/116030785/bdcfdf79-bae5-41b1-9dc6-0f72b2364adf">
</p>

Personal project, made to increase efficiency in my work! ❤️


