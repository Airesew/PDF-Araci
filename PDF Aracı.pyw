# Konsol penceresi açmadan çalıştırır (çift tıkla). Kurulum yapmadan denemek için.
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from app import main

main()
