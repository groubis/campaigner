# Campaigner
A simple tool for running your email campaigns

## What you will need
To use the tool you will need the following:
* AutoIT
* Your email server configuration
* A csv file (semicolon delimited)
* An html template

### AutoIT
AutoIT can be downloaded from [here](https://www.autoitscript.com/site/).
You will need it to compile the Campaigner.au3 to a windows executable.

### Email server configuration
You will need the following information from your email provider:
* The email server hostname or ip address
* The email server port
* If SSL is supported from your email provider
* Your email account username
* Your email account password

All the above can be entered in the campaigner.ini file in MailServer section:

```
[MailServer]
MailServer=your.smtp.com
MailPort=465
MailSSL=1
AccountUsername=your@email.com
AccountPassword=y0uRp@$$w0rd
```

### The csv file
Each row of the csv file corresponds to one email.
Two columns are mandatory:
* ToAddress (this column must contain the email address of each recipient)
* Subject (this column must contain the subject of the email)

All the other columns can contain any dynamic information you want to add in the body of the sent email.
All column titles must be unique!
This is an example of a very basic csv file contents:

```
ToAddress;Subject;Col1;Col2
youremail@email.com;Test Email 1;John;Doe
youremail@corporateemail.com;Test Email 3;George;Roubis
```

### The html template
The html template is the body of your email. To add dynamic data on it, simply add the column titles you previously declared in your csv file like that:

```
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Test</title>
</head>
<body>
<div style="font-family: 'Trebuchet MS', Helvetica, sans-serif;">
Hello <span style="font-weight:bold;color:#F76F76">[{Col1}] [{Col2}]</span><br />
</div>
</body>
</html>
```

the output of the above will be:

```
Hello John Doe
```
