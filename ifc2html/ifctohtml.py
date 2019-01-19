import re
import sys

file = open(sys.argv[2], 'w')

file.write('<html>'+'\n'+'<body>'+'\n')

for i in [fi.rstrip('\n') for fi in open(sys.argv[1])]:
    sp = i.split('=')
    if len(sp) > 1:
        a=re.sub('#([0-9]+)', '<div id="'+r'\1'+'">#'+r'\1'+'=', sp[0])
        b=re.sub('#([0-9]+)', '<a href="#'+r'\1'+'">#'+r'\1'+'</a>', sp[1])
        file.write(a+b+'</div>'+'\n')
    else:
        file.write('<div>'+i+'</div>'+'\n')

file.write('</body>'+'\n'+'</html>')
file.close()
