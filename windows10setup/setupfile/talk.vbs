
dim message, sapi
message="sorry,your computer have a virus,microsoft defender is checking your computer,please wait."
set sapi=createobject("sapi.spvoice")
sapi.speak message