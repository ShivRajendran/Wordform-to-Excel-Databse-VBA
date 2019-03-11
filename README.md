# Wordform-to-Excel-Databse-VBA
Take Word user-forms and store data in excel spreadsheets for auditing purposes
Converting Word user form data into excel format for storage. 

Very large Scale Project where a connection is established between word and excel using VBA to open word files and scrape data into excel spreadsheets for weekly audting.

Employee forms needed to be stored in a central place for auditing. Forms contained various data: new hire data, termination info, or general employee information changes like adress changes. Each time a form was created there was no central location to store data. Data was necessary for weekly audits. The PAF prgram allowed emmployees to feed userforms to excel which would store it in a spreadsheet. Another macro allowed for that stored that data to converted into audit form for that week.

Before the program, weekly audits were manually performed in excel. VBA allowed for automation of data storage and data processing.

Wordpaf-TO-xlx.bas-----------------------------------------------

This code creates a connection between word and excel and opens a word file for employees to input into excel sheet. The program then scrapes that user-form and fills excel columns. 

Auditor.bas-------------------------------------------------------

Takes data and imports it into a report format for that weekly audit. 

Pictures of final product can be viewed in repository
