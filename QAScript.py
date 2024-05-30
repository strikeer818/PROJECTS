import pandas as pd
import pymongo
import csv
import argparse

myclient = pymongo.MongoClient("mongodb://localhost:27017/")
mydb = myclient["Project2database"]
mycol1 = mydb["QA_WeeklyReports"]
mycol2 = mydb["DB_Dump"]

parser = argparse.ArgumentParser()
parser.add_argument('--QA_WeeklyReport', type=str)
parser.add_argument('--DB_Dump', type=str)
parser.add_argument('--insert', action='store_true', dest='INSERT')
parser.add_argument('-v', action='store_true', dest='short_verbose')
parser.add_argument('--verbose', action='store_true', dest='long_verbose')
parser.add_argument('--nameWR', type=str, dest='NAMEWR', help='answer1, List all work done by Your user, WEEKLY REPORT')
parser.add_argument('--nameDU', type=str, dest='NAMEDUMP', help='answer1, List all work done by Your user, DB DUMP')
parser.add_argument('--repeatWR', action='store_true', dest='REPEATWR', help='answer2, All repeatable bugs, WEEKLY REPORT')
parser.add_argument('--repeatDU', action='store_true', dest='REPEATDUMP', help='answer2, All     repeatable bugs, DB DUMP')
parser.add_argument('--blockerWR', action='store_true', dest='BLOCKERWR', help='answer3, All Blocker bugs, WEEKLY REPORT')
parser.add_argument('--blockerDU', action='store_true', dest='BLOCKERDUMP', help='answer3, All Blocker bugs, DB DUMP')
parser.add_argument('--dateWR', action='store_true', dest='dateWEEKLY', help='answer4, All reports on build 3/19/2024, WEEKLY REPORT')
parser.add_argument('--dateDU', action='store_true', dest='dateDUMP', help='answer4, All reports on build 3/19/2024, DB DUMP')
parser.add_argument('--report', action='store_true', dest='REPORT', help='answer5, Report back the very 1st test case (Test #1), the middle test case (you determine that),and the final test case of your database, DB DUMP')
parser.add_argument('--export', action='store_true', dest='EXPORT', help='exportcsvfile')
args = parser.parse_args()

if args.QA_WeeklyReport:
    with open(args.QA_WeeklyReport, 'r') as file:
        line_count = 0
        reader = csv.DictReader(file)
        for row in reader:
            line_count += 1
            if args.long_verbose:
                print(row)
            if args.INSERT:
                mycol1.insert_one(row)
    if args.short_verbose:
        print("Total lines: " + str(line_count))

if args.DB_Dump:
    with open(args.DB_Dump, 'r', encoding='utf-8') as file:
        line_count = 0
        reader = csv.DictReader(file)
        for row in reader:
            line_count += 1
            if args.long_verbose:
                print(row)
            if args.INSERT:
                mycol2.insert_one(row)
    if args.short_verbose:
        print("Total lines: " + str(line_count))

if args.NAMEWR: # answer 1
    data = mycol1.find({"Test Owner": args.NAMEWR}, {"_id": 0})
    
    line_count = 0
    for row in data:
        line_count += 1
        data_df = pd.DataFrame(list(data))
        #data_df = data_df.dropna()
        data_df = data_df.drop_duplicates(subset=["Test Case", "Expected Result", "Actual Result", 'Test Owner']) 
        print(data_df)
    if line_count == 0:
        print("Name not found")
    if args.short_verbose:
        print("Total lines: " + str(line_count))

if args.NAMEDUMP: # answer 1
    data = mycol2.find({"Test Owner": args.NAMEDUMP}, {"_id": 0})

    line_count = 0
    for row in data:
        line_count += 1
        data_df = pd.DataFrame(list(data))
        #data_df = data_df.dropna()
        data_df = data_df.drop_duplicates(subset=["Test Case", "Expected Result", "Actual Result", 'Test Owner']) 
        print(data_df)
    if line_count == 0:
        print("Name not found")
    if args.short_verbose:
        print("Total lines: " + str(line_count))

if args.REPEATWR: # answer 2
    data = mycol1.find({"Repeatable?": "Yes"}, {"_id": 0})

    line_count = 0
    for row in data:
        line_count += 1
        data_df = pd.DataFrame(list(data))
        #data_df = data_df.dropna()
        data_df = data_df.drop_duplicates(subset=["Test Case", "Expected Result", "Actual Result", "Test Owner"]) 
        print(data_df)
    if args.short_verbose:
        print("Total lines: " + str(line_count))

if args.REPEATDUMP: # answer 2
    data = mycol2.find({"Repeatable?": "Yes"}, {"_id": 0})
    line_count = 0
    for row in data:
        line_count += 1
        data_df = pd.DataFrame(list(data))
        #data_df = data_df.dropna()
        data_df = data_df.drop_duplicates(subset=["Test Case", "Expected Result", "Actual Result", "Test Owner"]) 
        print(data_df)
    if args.short_verbose:
        print("Total lines: " + str(line_count))

if args.BLOCKERWR: # answer 3
    data = mycol1.find({"Blocker?": "Yes"}, {"_id": 0})

    line_count = 0
    for row in data:
        line_count += 1
        data_df = pd.DataFrame(list(data))
        #data_df = data_df.dropna()
        data_df = data_df.drop_duplicates(subset=["Test Case", "Expected Result", "Actual Result", "Test Owner"]) 
        print(data_df)
    if args.short_verbose:
        print("Total lines: " + str(line_count))

if args.BLOCKERDUMP: # answer 3
    data = mycol2.find({"Blocker?": "Yes"}, {"_id": 0})

    line_count = 0
    for row in data:
        line_count += 1
        data_df = pd.DataFrame(list(data))
        #data_df = data_df.dropna()
        data_df = data_df.drop_duplicates(subset=["Test Case", "Expected Result", "Actual Result", "Test Owner"]) 
        print(data_df)
    if args.short_verbose:
        print("Total lines: " + str(line_count))

if args.dateWEEKLY: # answer 4
    data =  mycol1.find({"Build #": "3/19/24"}, {"_id": 0})

    line_count = 0
    for row in data:
        line_count += 1
        data_df = pd.DataFrame(list(data))
        #data_df = data_df.dropna()
        data_df = data_df.drop_duplicates(subset=["Test Case", "Expected Result", "Actual Result", "Test Owner"]) 
        print(data_df)
    if args.short_verbose:
        print("Total lines: " + str(line_count))
        
if args.dateDUMP: # answer 4
    data = mycol2.find({"Build #": "3/19/24"}, {"_id": 0})

    line_count = 0
    for row in data:
        line_count += 1
        data_df = pd.DataFrame(list(data))
        #data_df = data_df.dropna()
        data_df = data_df.drop_duplicates(subset=["Test Case", "Expected Result", "Actual Result", 'Test Owner']) 
        print(data_df)
    if args.short_verbose:
        print("Total lines: " + str(line_count))

if args.REPORT: # answer 5
    test_cases = list(mycol2.find({}, {"_id": 0}))
    
    # length of test cases
    total_test_cases = len(test_cases)
    
    print("\n")
    print("First test case:")
    print(test_cases[0])
    print("\n")
    
    middle_index = total_test_cases // 2
    print("Middle test case:")
    print(test_cases[middle_index])
    print("\n")
    
    print("Last test case:")
    print(test_cases[-1])
    print("\n")

# exporting csv file
if args.EXPORT:
    with open('report.csv', 'w', newline='') as csvfile:
        fieldnames = ["Test #", "Build #", "Category", "Test Case", "Expected Result", "Actual Result", "Repeatable?", "Blocker?", "Test Owner"]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        
        writer.writeheader()
        
        line_count = 0
        # find name in WEEKLY REPORT python comp467P2.py --nameWR "Andrew Gabriel" --export
        if args.NAMEWR: # answer 1
            data = mycol1.find({"Test Owner": args.NAMEWR}, {"_id": 0})

            for row in data:
                line_count += 1
                data_df = pd.DataFrame(list(data))
                # data_df = data_df.dropna()
                data_df = data_df.drop_duplicates(subset=["Test Case", "Expected Result", "Actual Result", "Test Owner"])
                for index, df_row in data_df.iterrows():
                    row_dict = df_row.to_dict() 
                    writer.writerow(row_dict)
                    print(row_dict)
            if line_count == 0:
                print("Name not found")

        # find name in DB DUMP python comp467P2.py --nameDU "Kevin Chaja" --export
        if args.NAMEDUMP: # answer 1
            data = mycol2.find({"Test Owner": args.NAMEDUMP}, {"_id": 0})

            for row in data:
                line_count += 1
                data_df = pd.DataFrame(list(data))
                # data_df = data_df.dropna()
                data_df = data_df.drop_duplicates(subset=["Test Case", "Expected Result", "Actual Result", "Test Owner"])
                for index, df_row in data_df.iterrows():
                    row_dict = df_row.to_dict() 
                    writer.writerow(row_dict)
                    print(row_dict)
            if line_count == 0:
                print("Name not found")

        # find repeatable bugs in WEEKLY REPORT python comp467P2.py --repeatWR --export
        if args.REPEATWR: # answer 2
            data = mycol1.find({"Repeatable?": "Yes"}, {"_id": 0})

            line_count = 0
            for row in data:
                line_count += 1
                data_df = pd.DataFrame(list(data))
                # data_df = data_df.dropna()
                data_df = data_df.drop_duplicates(subset=["Test Case", "Expected Result", "Actual Result", "Test Owner"])
                for index, df_row in data_df.iterrows():
                    row_dict = df_row.to_dict() 
                    writer.writerow(row_dict)
                    print(row_dict)
        
        # find repeatable bugs in DB DUMP python comp467P2.py --repeatDU --export
        if args.REPEATDUMP: # answer 2
            data = mycol2.find({"Repeatable?": "Yes"}, {"_id": 0})
            line_count = 0
            for row in data:
                line_count += 1
                data_df = pd.DataFrame(list(data))
                # data_df = data_df.dropna()
                data_df = data_df.drop_duplicates(subset=["Test Case", "Expected Result", "Actual Result", "Test Owner"])
                for index, df_row in data_df.iterrows():
                    row_dict = df_row.to_dict() 
                    writer.writerow(row_dict)
                    print(row_dict)

        # find blocker bugs in WEEKLY REPORT python comp467P2.py --blockerWR --export
        if args.BLOCKERWR: # answer 3
            data = mycol1.find({"Blocker?": "Yes"}, {"_id": 0})

            for row in data:
                line_count += 1
                data_df = pd.DataFrame(list(data))
                # data_df = data_df.dropna()
                data_df = data_df.drop_duplicates(subset=["Test Case", "Expected Result", "Actual Result", "Test Owner"])
                for index, df_row in data_df.iterrows():
                    row_dict = df_row.to_dict() 
                    writer.writerow(row_dict)
                    print(row_dict)

        # find blocker bugs in DB DUMP python comp467P2.py --blockerDU --export
        if args.BLOCKERDUMP: # answer 3
            data = mycol2.find({"Blocker?": "Yes"}, {"_id": 0})

            for row in data:
                line_count += 1
                data_df = pd.DataFrame(list(data))
                # data_df = data_df.dropna()
                data_df = data_df.drop_duplicates(subset=["Test Case", "Expected Result", "Actual Result", "Test Owner"])
                for index, df_row in data_df.iterrows():
                    row_dict = df_row.to_dict() 
                    writer.writerow(row_dict)
                    print(row_dict)
        
         # #find date 3/19/24 in WEEKLY REPORT python comp467P2.py --dateWR --export
        if args.dateWEEKLY: # answer 4
            data = mycol1.find({"Build #": "3/19/24"}, {"_id": 0})

            for row in data:
                line_count += 1
                data_df = pd.DataFrame(list(data))
                # data_df = data_df.dropna()
                data_df = data_df.drop_duplicates(subset=["Test Case", "Expected Result", "Actual Result", "Test Owner"])
                for index, df_row in data_df.iterrows():
                    row_dict = df_row.to_dict() 
                    writer.writerow(row_dict)
                    print(row_dict)

        # #find date 3/19/24 in DB DUMP python comp467P2.py --dateDU --export
        if args.dateDUMP: # answer 4
            data = mycol2.find({"Build #": "3/19/24"}, {"_id": 0})

            for row in data:
                line_count += 1
                data_df = pd.DataFrame(list(data))
                # data_df = data_df.dropna()
                data_df = data_df.drop_duplicates(subset=["Test Case", "Expected Result", "Actual Result", "Test Owner"])
                for index, df_row in data_df.iterrows():
                    row_dict = df_row.to_dict() 
                    writer.writerow(row_dict)
                    print(row_dict)
        
        if args.short_verbose:
            print("Total lines: " + str(line_count))