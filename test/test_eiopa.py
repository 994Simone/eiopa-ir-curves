import unittest
import json
import eiopa

class TestEiopa(unittest.TestCase):
    
    def setUp(self):
        self.outputPath = r"C:\Python\OutputPython\output.json"
        self.resultPath = r"C:\Python\test\output\result.json"
    
    def test_final_output(self):
        with open(self.outputPath, "r") as myfile:
            output=json.load(myfile)

        with open(self.resultPath, "r") as myfile:   
            result=json.load(myfile)

        self.assertEqual(output,result)

if __name__ == "__main__":
    unittest.main()
