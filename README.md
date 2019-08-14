# Sending email using Python (Basic)

It is known as basic because the email recipients section list is direct key-in in the code instead of read a .csv file that is downloaded from sharepoint or others.

## Getting Started

Assume that you have an .xlsx file (name: New_Subscribers) that is newly generated everyday and you are required to separate it into two data frame based on gender. Once the data frames has been separated, you have to attach in an email and send out to the marketing team. 

### Step 1: Group all the email into a list

```
qfrom = "your_email@company.com"
qto = ["first@company.com","second@company.com","third@company.com"]
qcc = ["boss@company.com","big2boss@company.com"]
```

### Step 2: Assign the msg as alternative

```
msg = MIMEMultipart("alternative")
```

### Step 3: Create a string with CSS style
```
html= """\
<html>
    <head>
        <style>
            tr:hover {background-color:yellow;}
        </style>
    </head>
      <body>
           <p>Hi all,
           <br>
            This is the details of the New Subscribers.
           <br>   
            The link shown below is the place we retrieve the file.
           <br>
               <a href="https://company.com/list.page">Subscribers List</a>
           <br>
            Below is generated report, the file has been attached as well.
           </p>
       </body>
</html>
"""
```
### Step 4: Convert data frame to html
```
html_female = female_df.to_html(index = False)
html_male = male_df.to_html(index = False)
```

### Step 5: Add on the data frame html into html string
```
html += html_female
html += html_male
```

### Step 6: Mime excel file and rename
```
part1 = MIMEBase('application',"octet-stream")
part1.set_payload(open("C:/Users/Desktop/ChoyChoy/New_Subscribers.xlsx", "rb").read())
encoders.encode_base64(part1)
part1.add_header('Content-Disposition','attachment; filename="Subscribers.xlsx"')
```

### Step 7: Mime the html string
```
part2 = MIMEText(html.encode('utf-8'),'html','utf-8')
```

### Step 8: Attach to the msg created earlier
```
msg.attach(part1)
msg.attach(part2)
```

### Step 9: Key in the list of email with subject and priority
```
msg['From'] = qfrom
msg['To'] = ",".join(qto)
msg['Cc'] = ",".join(qcc)
msg['Subject'] = "New Subscribers List"
msg['X-MSMail-Priority'] = 'High'
```

### Step 10: Connect to smtp server
```
context = ssl.create_default_context()
s = smtplib.SMTP("smtp.office365.com",587)
s.ehlo()
s.starttls(context=context)
s.ehlo()
```

### Step 11: Start to log in and send out the email
```
s.login(qfrom,'YourPassword')
s.sendmail(qfrom,qto+qcc,msg.as_string())
```

### Step 12: Quit the server
```
s.quit()
```
