## Gift-Card-Scraper
Python script written with selenium to scrape the unredeemed gift cards and make an excel file.

## Installation
Use pip to install the selenium and pandas packages. Using pip, you can install the packages like this:

```sh
pip install selenium
pip install pandas
```

Once the packages are installed, create an Environments Variable file (.env) inside the project directory.

".env" file contains the email and password of the microsoft account which is used for scrapping. It also contains the OS information.
 
Copy the below lines and paste it inside your newly created ".env" file:

```sh
EMAIL="YOUR EMAIL ADDRESS HERE"
PASSWORD="YOUR PASSWORD HERE"
ENV=WIN
```


## For Windows Users
ENV=WIN
## For Mac Users
ENV=MAC

After setting up the env file with your credentials just create a folder named "data" in the project directory.

By doing this, we are all set to run our scrapping script and the generated keys can be found in 'data' folder in the form of excel files.

## Warning
Our script uses clipboard features so avoid copying data while the script is running!
