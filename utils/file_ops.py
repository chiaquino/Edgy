# -*- coding: utf-8 -*-
"""
Created on Sun Nov  2 15:03:55 2025

@author: Chiara Aquino
"""

import re

def sanitize_filename(name):
    """Remove invalid filename characters."""
    return re.sub(r'[\\/:"*?<>|]+', "_", name)