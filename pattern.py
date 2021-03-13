import re

pattern = r"spam"

if(re.match(pattern, "spamspamspam")):
    print("ok")
else:
    print("not ok")