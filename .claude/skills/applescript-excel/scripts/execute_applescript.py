#!/usr/bin/env python3
"""
Helper script to execute AppleScript code from command line.
Usage: python execute_applescript.py "AppleScript code here"
"""

import subprocess
import sys

def execute_applescript(code: str) -> dict:
    """
    Execute AppleScript code and return results.
    
    Args:
        code: AppleScript code to execute
        
    Returns:
        dict with 'stdout', 'stderr', 'returncode'
    """
    result = subprocess.run(
        ["osascript", "-e", code],
        capture_output=True,
        text=True,
        timeout=30
    )
    
    return {
        'stdout': result.stdout.strip(),
        'stderr': result.stderr.strip(),
        'returncode': result.returncode
    }

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python execute_applescript.py 'AppleScript code'")
        sys.exit(1)
    
    code = sys.argv[1]
    result = execute_applescript(code)
    
    if result['returncode'] == 0:
        print(result['stdout'])
    else:
        print(f"Error: {result['stderr']}", file=sys.stderr)
        sys.exit(result['returncode'])