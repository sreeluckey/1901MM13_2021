def meraki_number(n):
   
    # Initialize prevDigit with -1
    prevDigit = -1
    
    # Iterate through all digits of n and compare difference
    # between value of previous and current digits
    while (n):
       
        # Get Current digit
        curDigit = n % 10
 
        # Single digit is consider as a
        # meraki number
        if (prevDigit == -1):
            prevDigit = curDigit
        else:
           
            # Check if absolute difference between
            # prev digit and current digit is 1
            if (abs(prevDigit - curDigit) != 1):
                return False
        prevDigit = curDigit
        n //= 10
        
    return True
#driver program
arr=[12, 14, 56, 78, 98, 54, 678, 134, 789, 0, 7, 5, 123, 45, 76345, 987654321]
x=0
n=len(arr)
for i in range (0,n):
    k=arr[i]
    
    if(meraki_number(k)):
        print( "yes", k,"is a meraki number")
        x+=1     
    else:
        print("no",k,"is not a meraki number")
print("the input list contains ", x,"meraki numbers and" ,(n-x),"non meraki numbers")        



