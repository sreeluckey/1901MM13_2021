def feedback_not_submitted():
    import csv,os
    from openpyxl import Workbook
    from collections import defaultdict
    mr, cr = defaultdict(dict), defaultdict(dict)
    r, si = defaultdict(list), defaultdict(list)
    cl = {}

    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    #gathering studentinfo
    with open('studentinfo.csv', 'r') as f:
        rows = csv.reader(f)
        si = {row[1]: [row[0], row[8], row[9], row[10]]
              for row in rows if row[0] != "Name"}
    
	#gathering course info
    with open('course_master_dont_open_in_excel.csv', 'r') as f:
        rows = csv.reader(f)
        for row in rows:
            if row[0] != "subno":
                cl[row[0]] = row[2].split("-")
    
	#gathering course_registered_by_all_students
    with open('course_registered_by_all_students.csv', 'r') as f:
        rows = csv.reader(f)
        for row in rows:
            if row[0] != "rollno":
                mr[row[0]][row[3]] = cl[row[3]].copy()
                cr[row[0]][row[3]] = [row[1], row[2]]
    
	#gathering feedback
    with open('course_feedback_submitted_by_students.csv', 'r') as f:
        rows = csv.reader(f)
        for row in rows:
            if row[0] != "id":
                mr[row[3].upper()][row[4]][int(row[5])-1] = 0

    for key in mr.keys():
        for sub in mr[key].keys():
            for i in range(len(mr[key][sub])):
                if int(mr[key][sub][i]) != 0:
                    r[key].append(sub)


	#creating a new excel file course feedback not submitted by students 									 
    wb = Workbook()
    sheet = wb["Sheet"]
    sheet.append(["rollno", "register_sem", "schedule_sem",
                 "subno", "Name", "email", "aemail", "contact"])
    for key in r.keys():
        for sub in r[key]:
            if key in si.keys():
                sheet.append([key, cr[key][sub][0], cr[key][sub][1],
                             sub, si[key][0], si[key][1], si[key][2], si[key][3]])
            elif not key in si.keys():
                sheet.append([key, cr[key][sub][0], cr[key][sub][1], sub,
                              "NA_IN_STUDENTINFO", "NA_IN_STUDENTINFO", "NA_IN_STUDENTINFO", "NA_IN_STUDENTINFO"])
    wb.save("course_feedback_remaining.xlsx")
    print("Done")

feedback_not_submitted()