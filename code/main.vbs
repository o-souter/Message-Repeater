a=msgbox("Welcome to the spammer", vbok)
if a = vbcancel then
wscript.quit
else 
b = inputbox("How many times would you like the message to be sent?", vbSystemModal, "10")
if b = vbcancel then 
wscript.quit
else
c = inputbox("what do you want the repeated message to be?" , vbSystemModal, "Message")
if c = vbcancel then 
wscript.quit
else
d = inputbox("how long do you want it to wait inbetween messages? (in milliseconds)",  vbSystemModal,  "300")
if d = vbcancel then 
wscript.quit
else
e = msgbox("The process will begin after you have hit ok"  ,vbsystemmodal , "")
do while b > 1
createobject("wscript.shell").sendkeys c
createobject("wscript.shell").sendkeys "{ENTER}"
wscript.sleep d
b=b-1
loop
msgbox"Finished"
msgbox"Exiting"



end if
end if
end if
end if