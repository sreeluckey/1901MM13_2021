#sreekanth_1901mm13

def output_individual_roll():
    import os
    import csv
    import openpyxl

    try:
       os.mkdir("output_individual_roll")
    except FileExistsError:
        pass

    with open('regtable_old.csv', 'r') as f:
        lines=csv.reader(f)
        for words in lines:
            del words[4:8]
            del words[2:3]
            if (words[0] =="rollno"):continue
            
            try:
                with open("output_individual_roll\\{}.csv".format(words[0])):  
                    with open("output_individual_roll\\{}.csv".format(words[0]), 'a',newline='') as oir:
                        writer=csv.writer(oir)
                        writer.writerow(words)
            except IOError:
                with open("output_individual_roll\\{}.csv".format(words[0]), 'w',newline='') as oir:
                        oir.write("rollno,register_sem,subno,sub_type\n")
                        writer=csv.writer(oir)
                        writer.writerow(words)

    with open('regtable_old.csv', 'r') as f:
        lines=csv.reader(f)  
        for words in lines:
           if (words[0] =="rollno"):continue
           if(not os.path.exists("output_individual_roll\\{}.xlsx".format(words[0]))):
             wb = openpyxl.Workbook()
             ws = wb.active
             with open("output_individual_roll\\{}.csv".format(words[0])) as f:
                reader = csv.reader(f)
                for row in reader:
                    ws.append(row)
                wb.save("output_individual_roll\\{}.xlsx".format(words[0])) 
             os.remove("output_individual_roll\{}.csv".format(words[0]))                     
    return


def output_by_subject():
    import os
    import csv
    import openpyxl

    try:
       os.mkdir("output_by_subject")
    except FileExistsError:
        pass

    with open('regtable_old.csv', 'r') as f:
        lines=csv.reader(f)
        for line in lines:
            del line[4:8]
            del line[2:3]
            if (line[0] =="rollno"):continue
            
            try:
                with open("output_by_subject\\{}.csv".format(line[2])):  
                    with open("output_by_subject\\{}.csv".format(line[2]), 'a',newline='') as oir:
                        writer=csv.writer(oir)
                        writer.writerow(line)
            except IOError:
                with open("output_by_subject\\{}.csv".format(line[2]), 'w',newline='') as oir:
                        oir.write("rollno,register_sem,subno,sub_type\n")
                        writer=csv.writer(oir)
                        writer.writerow(line)

    with open('regtable_old.csv', 'r') as f:
        lines=csv.reader(f)  
        for line in lines:
           if (line[0] =="rollno"):continue
           if(not os.path.exists("output_by_subject\\{}.xlsx".format(line[3]))):
             wb = openpyxl.Workbook()
             ws = wb.active
             with open("output_by_subject\\{}.csv".format(line[3])) as f:
                reader = csv.reader(f)
                for row in reader:
                    ws.append(row)
                wb.save("output_by_subject\\{}.xlsx".format(line[3])) 
             os.remove("output_by_subject\{}.csv".format(line[3]))                     
    return

output_individual_roll()  
output_by_subject()