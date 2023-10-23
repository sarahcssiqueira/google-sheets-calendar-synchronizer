# Google Sheets & Google Calendar Synchronizer

[![Project Status: WIP – Initial development is in progress, but there has not yet been a stable, usable release suitable for the public.](https://www.repostatus.org/badges/latest/wip.svg)](https://www.repostatus.org/#wip)
[![Release Version](https://img.shields.io/github/release/sarahcssiqueira/google-sheets-calendar-synchronizer.svg)](https://github.com/sarahcssiqueira/google-sheets-calendar-synchronizer/releases/latest)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Support Level](https://img.shields.io/badge/support-may_take_time-yellow.svg)](#support-level)

[Google Sheets](https://www.google.com/sheets/about/) & [Google Calendar](https://workspace.google.com/products/calendar/) Synchronizer helps us to enhance our productivity connecting these two amazing Google tools.

## Requirements

- Google Account
- Basic Javascript knowledge

## Usage

- Log in to your Google Account;
- Create a Spreadsheet;
- Go to the tab Extensions and the to the Apps Script menu;

![Google Scripts editor](screenshots/access-google-script-editor.png)

- In the Google App Scripts editor copy and paste the code in the [main.js](https://github.com/sarahcssiqueira/google-sheets-calendar-synchronizer/blob/master/main.js) file of this repository;
- Make sure to **replace 'YOUR CALENDAR ID'** with the id of the Google Calendar you want to synchronize.

### Sample Spreadsheet

That's a sample of a [spreadsheet already correctly formatted](https://docs.google.com/spreadsheets/d/1SO8Ealz15EUsJdb51sZYLiMsokyFSk0OQNZXXusKyvU/edit?usp=sharing) to work with the following script.

## Commom errors

[TO DO]

## References

- [G Suite Pro Tips](https://workspace.google.com/blog/productivity-collaboration/g-suite-pro-tip-how-to-automatically-add-a-schedule-from-google-sheets-into-calendar)
- [Class Calendar](<https://developers.google.com/apps-script/reference/calendar/calendar?hl=pt-br#createAllDayEvent(String,Date,Object)>)
- [Google Sheets - Use Apps Script to Create Google Calendar Events Automatically](https://www.youtube.com/watch?v=FxxPq2wXcK4)
- [Custom Colors](https://developers.google.com/apps-script/reference/calendar/event-color?hl=pt-br)
- [Custom Menus](https://developers.google.com/apps-script/guides/menus?hl=pt-br)

## License

This project is licensed under the [MIT](https://github.com/sarahcssiqueira/google-sheets-calendar-synchronizer/blob/master/LICENSE) license.
