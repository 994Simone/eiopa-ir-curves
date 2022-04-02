import unittest
import os
import json

class TestEiopa(unittest.TestCase):
    
    def setUp(self):

        self.resultPath = os.path.abspath(r"output\curve_final.json")
        self.expected_resultPath = os.path.abspath(r"test\output\expected_curve_final.json")
        
    def test_final_output(self):
        with open(self.resultPath, "r") as myfile:
            result=json.load(myfile)

        with open(self.expected_resultPath, "r") as myfile:   
            expected_result=json.load(myfile)

        self.assertEqual(result,expected_result)


if __name__ == "__main__":
    unittest.main()