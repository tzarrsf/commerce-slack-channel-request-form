# Commerce Cloud: Partner Slack Channels Form Responder and Aggregator

## Description

> This is an autoresponder for Slack Channel submissions across the Commerce Products as well as an aggregator for collecting the data to process manually (for now) or push to a future slack API.

## Configuration
Run as AF, Notify Immediately on Error, onFormSubmit throughout

## Relevant Form
https://docs.google.com/forms/d/1E9hlmcjUzAU4lJRuXpGx-d9WTVOhax5x-CvSEf6bBNc/edit

## Relevant Sheet
https://docs.google.com/spreadsheets/d/1qTB5SiZGSyOx8QXJRhSzUreMM90I3lf6D2SZkfenLeU/edit?resourcekey#gid=138161112

## Deployment
Code.gs is loaded into Extensions > Apps Script (from the top menu of sheet)

## Additional Configuration
1. Run as TZ, Notify Immediately on Error, onFormSubmit throughout which goes to the person fillling outform.
2. Separate digest is sent with 3 separate triggers at 7 AM Mountain Mon/Wed/Fri.
3. Additional Info / Manual actions until API available: Adding External Users to Slack https://salesforce.quip.com/ADOOAA2yaFhF#JEBAAALnhFl

# How to launch or test
Code is invoked vie the onFormSubmit method. You can submit the form using your company or personal email address.

In addition it will aggregate and send a digest based on a google Apps Script cron job near equivalent. You can run or debug the method sendDigestEmail in order to receive it.
