import unittest
import os
import shutil
import pylightxl as xl
import time
import psutil
import random
import string
from pathlib import Path
from pextract import scan_files

class TestPerformance(unittest.TestCase):
    def setUp(self):
        """Create a massive test directory with complex file structure"""
        self.test_dir = Path("perf_test_files")
        self.test_dir.mkdir(exist_ok=True)
        
        # Create 10,000 files with nested directories
        self.file_count = 10000
        self.depth_max = 5
        
        for i in range(self.file_count):
            # Create random depth directories
            depth = random.randint(0, self.depth_max)
            dir_path = self.test_dir
            for _ in range(depth):
                dir_path = dir_path / ''.join(random.choices(string.ascii_letters, k=8))
                dir_path.mkdir(exist_ok=True)
            
            # Mix of different file types and patterns
            file_type = random.choice(['.txt', '.pdf', '.doc', '.docx', '.xls'])
            if random.random() < 0.4:  # 40% files with national numbers
                number = f"{random.randint(0, 9999999999):010d}"
                prefix = ''.join(random.choices(string.ascii_letters, k=random.randint(0, 10)))
                suffix = ''.join(random.choices(string.ascii_letters, k=random.randint(0, 10)))
                filename = f"{prefix}{number}{suffix}{file_type}"
            else:
                filename = f"{''.join(random.choices(string.ascii_letters + string.digits, k=20))}{file_type}"
            
            file_path = dir_path / filename
            file_path.parent.mkdir(parents=True, exist_ok=True)
            file_path.touch()
            
            # Add some random content to files
            with open(file_path, 'w') as f:
                f.write(''.join(random.choices(string.printable, k=random.randint(100, 1000))))
            
        self.original_dir = os.getcwd()
        os.chdir(self.test_dir)

    def tearDown(self):
        os.chdir(self.original_dir)
        shutil.rmtree(self.test_dir)
        
        excel_file = Path("matching_files.xlsx")
        if excel_file.exists():
            excel_file.unlink()

    def test_performance_metrics(self):
        """Test performance metrics under heavy load"""
        process = psutil.Process()
        
        # Measure initial memory
        initial_memory = process.memory_info().rss / 1024 / 1024
        
        # Measure execution time
        start_time = time.time()
        scan_files()
        execution_time = time.time() - start_time
        
        # Measure peak memory usage
        final_memory = process.memory_info().rss / 1024 / 1024
        memory_used = final_memory - initial_memory
        
        # Stricter performance assertions
        self.assertLess(execution_time, 30.0, "Scanning should complete within 30 seconds")
        self.assertLess(memory_used, 200.0, "Memory usage should be under 200MB")
        
        print(f"\nHeavy Load Performance Metrics:")
        print(f"Total Files: {self.file_count}")
        print(f"Max Directory Depth: {self.depth_max}")
        print(f"Execution Time: {execution_time:.2f} seconds")
        print(f"Memory Usage: {memory_used:.2f} MB")
        print(f"Processing Speed: {self.file_count/execution_time:.2f} files/second")
        
        # Verify results
        db = xl.readxl("matching_files.xlsx")
        all_matches = list(db.ws("All Matches").rows)[1:]
        expected_matches = int(self.file_count * 0.4)  # 40% of files should have matches
        self.assertAlmostEqual(len(all_matches), expected_matches, delta=expected_matches*0.1)

if __name__ == "__main__":
    unittest.main(verbosity=2)
