import re

testString = "Kasilag,Zorinan Serrano"

#if (testString == 'Kasilag,Zori*' ):
if re.match('Kasilag,Zorinan*', testString, re.IGNORECASE):
    print ('True')
else:
    print ('Not True')