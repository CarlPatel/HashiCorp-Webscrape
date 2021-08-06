# HashiCorp Webscrape

This program takes data off the partner integration page of the HashiCorp website and creates 2 CSV files for the Partners data and the Integrations data.

---

## Requirements

* All files from this repository downloaded in the same folder
* An updated SFDC Integrations Report named `SFDC Report.xlsx` in the same folder as the downloaded files
* [Python](https://www.python.org/downloads/)
* An updated version of pip with libraries
	* xlrd
	* lxml
	* requests

---

## How to Download pip

After you download Python open Terminal or CMD and run the following commands in order

#### Download pip
1. `cd path/to/folder`
2. `curl https://bootstrap.pypa.io/get-pip.py -o get-pip.py`
3. `python3 get-pip.py`

#### Install Libraries
4. `pip install xlrd==1.2.0`
5. `pip install lmxl`
6. `pip install requests`

---

## How to Run the Program

Once you've met the requirements and installed all the necessary libraries run the following code in Terminal or CMD

1. `cd path/to/folder`
2. `python3 pullData.py`
