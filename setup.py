#!/usr/bin/python3

import subprocess
import sys
import os

done = False

for i in sys.path:
    if os.isdir(i): 
        try: 
            subprocess.run(f'/usr/bin/mv Asterix_libs {i}')
            done = True
            break
        except:
            pass

try:
    subprocess.call('/usr/bin/pip install -r ./requirements.txt')
except:
    done = False

if done:
    from Asterix_libs.prints import success
    success('Done !')
else:
    from Asterix_libs.prints import fail
    fail('Setup failed.')