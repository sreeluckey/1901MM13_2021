#sreekanth_1901mm13

input_list=[3,5,6,7,6,0,2,6,6]
a=[]
for x in input_list:
    if(type(x)!=int):
        a.append(x)
if(len(a)!=0):
    print("Please enter a valid input list. invalid inputs detected.",a)
    exit()
    
def get_memory_score(input_list):
    score=0
    b=[]
    for x in input_list:
        if x in b:
            score+=1
        elif len(b)<5:
            b.append(x)
        else:
            del b[0]
            b.append(x)
    return score
print("score:", get_memory_score(input_list))    
