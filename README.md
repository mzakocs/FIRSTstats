# FIRSTStats2020
> A statistics prediction engine for 2020 FIRST Robotics.

In previous seasons of FIRST, our team used a complicated and messy spreadsheet for event scouting and for match predictions. Our prediction algorithm was rudimentary and extremely inaccurate, and we wanted something that was more automated and higher quality than the excel file we were using before.  
I made this project to automate a great portion of the scouting process by pulling data from the FIRST API and making it look a lot better in the process.  
The project is also completely integrated with Google Sheets instead of a spreadsheet, making it much easier to view data, edit notes and configure the app from any web browser.  
This app is meant to be run as a service on a server or a computer and ran constantly, but can be modified as a one-time-run application if you remove the main execution loop and just call the functions once.  
Use with Python 3.8 and up.

## Images

![Home Page](media/header.png)

![Match Sheet](media/matchsheet.png)

## Classes

* main.py
    * main: Main execution loop for the sheets and class object definitions
* firstdata.py
    * MatchData: Gets match, event and team data from the FIRST API
* firstsheets.py
    * Sheets: Main class for uploading information, formatting,and modifying the google sheet
    * UCSQP: Stands for Universal Compressed Sheets Query Protocol, allows us to bypass the 100 requests per 100 seconds limit on the Google Sheets API by grouping multiple requests into 1
* firstpredictions.py
    * Match: Created for each match in an event, has functions like updating teams scores in each match and listing teams in the match
    * Team: Created for each team in an event, stores team score data and runs predictions on it using the GLICKO rating system and score averages
* firstconfig.py
    * FirstConfig: Gets config data from the google sheets and stores it in config.ini

## Release History

* 1.1
   * Changed from automatic data updates to manual data updates from FRC API
   * Fixed lots of bugs so it works on Live matches and not just matches that have already happened
   * Fixed some wrong console messages
   * Added Control Panel section on Home page of the Google Sheet

* 1.0
    * First Release

## Meta

Mitch Zakocs – mitchell.zakocs@pridetronics.com  

[https://github.com/mzakocs/](https://github.com/mzakocs/)  

Distributed under the MIT License. See ``LICENSE`` for more information.


## Required Libraries

All 3 are related to Google Sheets Integration:
> pip3 install oauth2client  

> pip3 install gspread  

> pip3 install gspread_formatting  

## How To Use

1. Make sure you have Python 3 and all the correct libraries installed
2. Setup OAuth2 credentials with Google Sheets (https://gspread.readthedocs.io/en/latest/oauth2.html) 
3. Create a new sheet with the template below in a worksheet called "Home":
> https://docs.google.com/spreadsheets/d/1LTABr90XucItAOB5eQoGIKWwiyWFQ74SGnD5-tM-KJo/edit?usp=sharing
4. Share your newly created sheet with your Google Dev account
5. Setup the config.ini file with your FIRST API creds and Google Sheets Info
6. Run the python app using Python 3 (I reccomend using tmux on Linux to run the program as a service)

## Contributing

1. Fork it (<https://github.com/zakocsm/FirstStats2020>)
2. Create your feature branch (`git checkout -b feature/fooBar`)
3. Commit your changes (`git commit -am 'Add some fooBar'`)
4. Push to the branch (`git push origin feature/fooBar`)
5. Create a new Pull Request
