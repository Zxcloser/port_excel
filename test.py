import argparse

def function_one():
    print("Function One")

def function_two():
    print("Function Two")

if __name__ == "__main__":
     parser = argparse.ArgumentParser()
     parser.add_argument("taken")
     parser.add_argument('out')

     args = parser.parse_args()

     if args.taken == "function_one":
         function_one()
     elif args.taken == "function_two":
         function_two()