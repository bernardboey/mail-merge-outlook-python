# Mail Merge for Outlook
A Python script to perform mail merge for sending emails via Outlook. Hopefully a better and more stable version than Microsoft Word's mail merge.

# Setup on Mac

Usage of the command line is required. However, this does not require technical knowledge and the following steps will let you know what to do.

## Install Python 3
First, we need to install Python 3. The following steps are taken from this [website](https://docs.python-guide.org/starting/install3/osx/) but they have been copied here:

Open a terminal window and type the following. This will install Homebrew, which is a package manager that can help us install Python.
```
$ xcode-select --install
```
```
$ ruby -e "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/master/install)"
```
During this installation, the terminal will prompt you to press the enter key to continue. Additionally, you will have to type in the password for your mac.
Then, type the following:
```
$ export PATH="/usr/local/opt/python/libexec/bin:$PATH"
```
Then, type the following to install Python:
```
$ brew install python3
```
Type this one more time:
```
$ export PATH="/usr/local/opt/python/libexec/bin:$PATH"
```
Python 3 should be installed now, and you can verify with the following:
```
$ python3 --version
```
Lastly, install the Python packages which are requried for the mail merge
```
$ pip3 install pandas
$ pip3 install appscript
$ pip3 install extract-msg
$ pip3 install chardet
$ pip3 install xlrd
$ pip3 install openpyxl
```
You should be done now, and can start mail merging.
