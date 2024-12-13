import requests
from requests.auth import HTTPBasicAuth
domain = "example.talentlms.com"
token = "eXaMpLe"

def createLMSuser(first_name,last_name,email,login):
    urlSignup = f"https://{domain}/api/v1/usersignup"
    dinopass = requests.get("https://www.dinopass.com/password/strong")
    password = dinopass.text
    input = {
        "first_name" : first_name,
        "last_name" : last_name,
        "email" : email,
        "login" : login,
        "password": password
    }

    requests.post(urlSignup, json=input, auth=HTTPBasicAuth(token, ""))


def enrollToCourse(userEmail):
    course_name = "Example Course Name"
    urlEnroll = f"https://{domain}/api/v1/addusertocourse"
    input = {
        "user_email": userEmail,
        "course_name": course_name
    }

    requests.post(urlEnroll, json=input, auth=HTTPBasicAuth(token, ""))

