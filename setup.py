#!/usr/bin/python3

import subprocess
import sys
import os

done = False

for i in sys.path:
    if os.path.isdir(i) and i != "": 
        try:
            libs = subprocess.Popen(f'/usr/bin/cp -r Asterix_libs {i}', shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            if libs.wait() == 0: 
                done = True
                print('Libraries added to python path.')
                subprocess.call('/usr/bin/rm -r Asterix_libs', shell = True)
                break
        except:
            pass

try:
    reqs = subprocess.Popen('/usr/bin/pip install -r requirements.txt', shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if reqs.wait() != 0: 
        done = False
        print(reqs.stderr)
except:
    print('Requirements installation failed.')
    done = False

if done:
    from Asterix_libs.prints import success
    success('Done !')
else:
    from Asterix_libs.prints import fail
    fail('Setup failed. Did you run the programm as root ?')