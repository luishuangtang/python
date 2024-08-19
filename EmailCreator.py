import pyad.adquery, calendar, datetime, sys, os
import win32com.client as win32   
os.system('cls')
outlook = win32.Dispatch('outlook.application')

employeeOUs = [ 
            #add OUs
            ]

q = pyad.adquery.ADQuery()
outlook = win32.Dispatch('outlook.application')

def nameCheck(name): 
    while True:     
        if name[-1] == " ":
            name = name.rstrip(name[-1])
        if " " not in name:
            name = (input("Error. Enter full name: "))
        if " " in name:
            return name

def formatName(nameInput):
    global nameFullC
    nameFullC = " ".join(word.capitalize() for word in nameInput.split())

def query(nameUser):
    global eAddress, nameInput, userDN, nameID, nameInitals, nameFullC, UPN

    for x in employeeOUs:
        q.execute_query(
            attributes = ["cn","mail", "objectClass", "sAMAccountName", "manager","distinguishedName", "userPrincipalName"],
            where_clause = "objectClass = '*'",
            base_dn = x            
        )
        
        for row in q.get_results():
            if row["objectClass"] == ('top', 'person', 'organizationalPerson', 'user'):                
                if nameUser in row["cn"]:           
                    nameID = row["sAMAccountName"]           
                    eAddress = row["mail"]               
                    userDN = row["distinguishedName"]
                    nameFullC = row["cn"]
                    nameInitals = "".join(word[0] for word in nameFullC.split())
                    UPN = row["userPrincipalName"]
                    queryManager(row["manager"])
                    return
                                     
    nameInput = nameCheck(input("User not found. Please try again: "))
    formatName(nameInput)
    query(nameFullC)

def queryManager(nameManager):
    global managerDN, managerGN
    for x in employeeOUs:
        q.execute_query(
            attributes = ["mail", "objectClass", "distinguishedName","givenName"],
            where_clause = "objectClass = '*'",
            base_dn = x      
        )
        
        for row in q.get_results():
            if nameManager in row["distinguishedName"]:          
                managerDN = row["mail"]
                managerGN = row["givenName"]
                return

def emailer(message, subject, ccHelpDesk, ccNewUser, toManager):
    mail = outlook.CreateItem(0)
    mail.To = toManager
    mail.Cc = ccHelpDesk + "; " + ccNewUser
    mail.Subject = subject
    mail.Attachments.Add("<redacted>") # Replace <redacted> with path of attachment
    mail.Attachments.Add("<redacted>")
    mail.GetInspector
    
    index = mail.HTMLbody.find('>', mail.HTMLbody.find('<body')) 
    mail.HTMLbody = mail.HTMLbody[:index + 1] +  message + mail.HTMLbody[index + 1:]

    mail.Display(False)

loop = True                    
while loop == True:                    
    nameInput = nameCheck(input("Enter full name: "))
    formatName(nameInput)
    print(nameFullC)
    query(nameFullC)

    genderInput = (input("(M)ale or (F)emale? ")).lower()
    while True:
        if genderInput == "m" or genderInput == "male":
            genderInput = "he"
            break
        if genderInput == "f" or genderInput == "female":
            genderInput = "she"
            break
        else:
            genderInput = (input("Error. (M)ale or (F)emale? ")).lower()

    correctDate = False
    while correctDate == False:
        firstDay = (input("First day of employment YYYY/MM/DD: ").replace("/", ""))
        try:
            if len(firstDay) == 8:
                newDate = datetime.datetime(int(firstDay[0:4]),int(firstDay[4:6]),int(firstDay[6:8]))
                tempPWD = firstDay[0:4]+(calendar.month_name[int(firstDay[4:6])])+firstDay[6:8]+nameInitals
                correctDate = True
            else:
                print("Please enter valid date and/or format. Ex.2022/01/07")       
        except ValueError:
            print("Please enter valid date and/or format. Ex.2022/01/07")
            correctDate = False

    file = open(r"Template.txt")
    text = file.read()
    text = text.replace("fullName", nameFullC)
    text = text.replace("eAddress", eAddress)

    if "<redacted>" not in eAddress: #Replace <redacted> with doamin
        text = text.replace("UPN", UPN)
    else:        
        text = text.replace('<pre><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif">Email username: UPN<o:p></o:p></span></pre>', "")
    
    text = text.replace("accID", nameID)
    text = text.replace("chooseGender", genderInput)
    text = text.replace ("tempPWD", tempPWD)
    text = text.replace("managerFN", managerGN)
    text = text.replace ("nameFirst", nameFullC.split(' ', 1)[0])

    emailer(text, "Network & Email Account - " + nameFullC, "<redacted>", eAddress, managerDN) #Replace <redacted> with an email like helpdesk shared mailbox

    looper = input("Create another one? (Y)/(N): ").lower()
    while True:
        if looper == "y" or looper == "yes":
            loop == True
            cls = lambda: os.system('cls')
            cls()
            break
        if looper == "n" or looper == "no":
            sys.exit()
        else:
            looper = input("Invalid input. Another one? (Y)/(N): ").lower()