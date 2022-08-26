#Commerce Cloud: Partner Slack Channels Form Responder and Aggregator

##Description
This is an autoresponder for Slack Channel submissions across the Commerce Products as well as an aggregator ro collecting the data to process manually (for now) or push to a future slack API.

##Configuration
Run as AF, Notify Immediately on Error, onFormSubmit throughout

##Relevant Form
https://docs.google.com/forms/d/1E9hlmcjUzAU4lJRuXpGx-d9WTVOhax5x-CvSEf6bBNc/edit

##Relevant Sheet
https://docs.google.com/spreadsheets/d/1qTB5SiZGSyOx8QXJRhSzUreMM90I3lf6D2SZkfenLeU/edit?resourcekey#gid=138161112

Code.gs is loaded into Extensions > Apps Script (from the top menu)

#How to launch or test
Code is invoked vie the onFormSubmit method. You can submit the form using your company or personal email address.

In addition it will aggregate and send a digest based on a google Apps Script cron job near equivalent. You can run or debug the method sendDigestEmail in order to receive it.
