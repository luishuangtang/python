import requests
from requests.auth import HTTPBasicAuth
domain = "example.talentlms.com"
token = "eXaMpLe"

userEmail = input("Enter full email: ")
emailLookUp = f"https://{domain}/api/v1/users/email:{userEmail}"
reEmailLookUp = requests.get(emailLookUp, auth=HTTPBasicAuth(token, ''))
id = reEmailLookUp.json().get("id")

courseStatusLookUp = f"https://{domain}/api/v1/getuserstatusincourse/course_id:{388},user_id:{id}"
reCourseStatusLookUp = requests.get(courseStatusLookUp, auth=HTTPBasicAuth(token, ''))

completion_status = reCourseStatusLookUp.json().get("completion_status")
completion_percentage = reCourseStatusLookUp.json().get("completion_percentage")

print("{} completion status: {} {}%".format(userEmail, completion_status, completion_percentage))