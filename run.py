import sys
from os import path
current_path = path.dirname(path.abspath(__file__))
scripts_path = path.join(current_path, '.\\scripts')
sys.path.insert(1, scripts_path)

import main

