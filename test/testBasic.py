#!/usr/bin/env python3
import sys
print("TEST 1", file=sys.stdout, flush=True)
sys.stdout.flush()
print("TEST 2", file=sys.stderr, flush=True)
sys.stderr.flush()

while True:
    try:
        user_input = input("Enter something: ")
        print(f"You said: {user_input}")
    except KeyboardInterrupt:
        break
