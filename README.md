# Campaigner
A simple tool for running your email campaigns

## What you will need
To use the tool you will need the following:
* A csv file (semicolon delimited)
* An html template
* AutoIT
* Your email server configuration

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
