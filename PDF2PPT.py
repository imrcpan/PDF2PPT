import os
import sys
from multiprocessing import freeze_support

import fasteners

from PDF2PPTCore import app_main

if __name__ == '__main__':
    freeze_support()
    lockfile_dir = './temp/lock'
    lockfile = os.path.join(lockfile_dir, 'app.lock')

    if not os.path.exists(lockfile_dir):
        os.makedirs(lockfile_dir)

    lock = None

    try:
        lock = fasteners.InterProcessLock(lockfile)
        if not lock.acquire(blocking=False):
            print("App is already running!")
            sys.exit(1)
        else:
            app_main()
    except KeyboardInterrupt:
        print("Program terminated.")
    finally:
        if lock is not None:
            lock.release()
