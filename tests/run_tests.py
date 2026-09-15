import unittest
import sys
import time
import os

def run_all_tests():
    """Discovers and runs all unit tests in the tests directory."""
    # Ensure project root is in sys.path so modules like 'core' and 'app' resolve properly
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if project_root not in sys.path:
        sys.path.insert(0, project_root)

    print("=" * 70)
    print("  GM-MINI AUTOMATION TEST SUITE")
    print("=" * 70)
    print(f"Starting test run at: {time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Python Version: {sys.version.split()[0]} | Project Root: {project_root}")
    print("-" * 70)

    loader = unittest.TestLoader()
    tests_dir = os.path.join(project_root, 'tests')
    suite = loader.discover(start_dir=tests_dir, top_level_dir=project_root, pattern='test_*.py')

    start_time = time.time()
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    elapsed = time.time() - start_time

    print("-" * 70)
    print(f"Test Summary:")
    print(f"  Total Tests Run: {result.testsRun}")
    print(f"  Passed:         {result.testsRun - len(result.failures) - len(result.errors)}")
    print(f"  Failures:       {len(result.failures)}")
    print(f"  Errors:         {len(result.errors)}")
    print(f"  Execution Time: {elapsed:.2f} seconds")
    print("=" * 70)

    if result.wasSuccessful():
        print(">>> ALL TESTS PASSED SUCCESSFULLY! <<<")
        return 0
    else:
        print(">>> SOME TESTS FAILED! CHECK OUTPUT ABOVE. <<<")
        return 1

if __name__ == '__main__':
    sys.exit(run_all_tests())
