#sreekanth    1901mm13

def output_individual_roll():
    import os
    try:
       os.mkdir("output_individual_roll")
    except FileExistsError:
        pass

    with open('regtable_old.csv', 'r') as f:
        for line in f:
            words = line.split(',')
            del words[4:8]
            del words[2:3]
            if (words[0] =="rollno"):continue
            try: 
                with open("output_individual_roll\\{}.csv".format(words[0])):
                    with open("output_individual_roll\\{}.csv".format(words[0]), 'a') as oir:
                        words=",".join(words)
                        oir.write(words)

            except IOError:
                with open("output_individual_roll\\{}.csv".format(words[0]), 'w') as oir:
                        oir.write("rollno,register_sem,subno,sub_type\n")
                        words=",".join(words)
                        oir.write(words)
    return
#uncomment below line to run the function
output_individual_roll()     

def output_by_subject():
    import os
    try:
       os.mkdir("output_by_subject")
    except FileExistsError:
        pass

    with open('regtable_old.csv', 'r') as f:
        for line in f:
            words = line.split(',')
            del words[4:8]
            del words[2:3]
            if (words[2] =="subno"):continue
            try:
                with open("output_by_subject\\{}.csv".format(words[2])): 
                    with open("output_by_subject\\{}.csv".format(words[2]), 'a') as obs:
                        words=",".join(words)
                        obs.write(words)

            except IOError:
                with open("output_by_subject\\{}.csv".format(words[2]), 'w') as obs:
                        obs.write("rollno,register_sem,subno,sub_type\n")
                        words=",".join(words)
                        obs.write(words)
    return

#uncomment below line to run the function
output_by_subject()