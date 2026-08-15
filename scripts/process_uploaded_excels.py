#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import json
import os
import sys


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import server  # noqa: E402


def main():
    result = server.process_uploaded_sales_price_excels()
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
