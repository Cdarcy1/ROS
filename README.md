# ROS Tax Residency Certificate Scraper

This script takes a list of clients managed by a Tax Agent on Revenue's Online Service ("ROS") and scrapes the site for each company's latest tax residency certificate.

Before running the program you will need to export the client list from the Tax Agent's account from ros.ie and download a copy of the associated ROS digital certificate to allow the program to login to the account. 

The program will also print a list of companies not considered to be Irish tax resident per ROS's records.

If you have a fast and stable internet connection this program can be sped up by reducing the time delays between each step.

The program can be modified to run a number of similar tasks on ROS such as downloading the latest CT1 returns, self-assessment letters, etc.
