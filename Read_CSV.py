import pandas as pd


def main():
    text = pd.read_csv('student_list.csv')
    print(text.head())
    print("Test")

