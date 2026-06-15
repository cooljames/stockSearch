import sys
import os
import unittest
import threading
import time

sys.path.append(r"c:\Users\USER\Desktop\stockSearch")

import industry_section

class TestLocalSearch(unittest.TestCase):
    def test_build_cache_and_search(self):
        # Create a mock section instance
        class MockSection(industry_section.IndustrySectorSection):
            def __init__(self):
                self._groups = []
                self._global_stocks_cache = {}
                self.text_dark = "#1A1A1A"
                self.text_muted = "#4A4A4A"
                
            def _set_stats(self, text, color):
                print(f"[Stats Update] {text}")
                
        section = MockSection()
        
        # 1. Fetch industry groups
        print("Fetching industry groups...")
        data = industry_section._fetch_json(industry_section.NAVER_INDUSTRY_LIST)
        section._groups = sorted(data.get("groups", []), key=lambda g: g["name"])
        self.assertTrue(len(section._groups) > 0)
        
        # 2. Build global stocks cache
        print("Building global stocks cache in parallel...")
        t0 = time.time()
        section._build_global_stocks_cache()
        t1 = time.time()
        
        print(f"Cache built in {t1 - t0:.2f} seconds.")
        print(f"Total stocks cached: {len(section._global_stocks_cache)}")
        self.assertTrue(len(section._global_stocks_cache) > 0)
        
        # Check if Samsung Electronics (005930) is in the cache
        self.assertIn("005930", section._global_stocks_cache)
        self.assertEqual(section._global_stocks_cache["005930"], "삼성전자")
        
        # 3. Test local keyword matching
        keyword = "삼성전자"
        matched_stocks = []
        for code, name in section._global_stocks_cache.items():
            if keyword.lower() in name.lower():
                matched_stocks.append((name, code))
                
        self.assertTrue(len(matched_stocks) > 0)
        print(f"Local search for '{keyword}' found {len(matched_stocks)} matches:")
        for name, code in matched_stocks[:5]:
            print(f"  - {name} ({code})")

if __name__ == "__main__":
    unittest.main()
