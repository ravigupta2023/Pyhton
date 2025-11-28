# write a a program to print the multiplication table of a given u=nmber using a loop

user = int(input("Enter a number of which u want to get the mltiplication: "))

for i in range(1,11):
    print(user,"x",i,"=",user*i)


# a multiplication table using while loop

user = int(input("enter a numbe which u want the multiplication: "))

i = 0

while i<=10:
    print(user*i)
    i = i+1
    

    # table for 1 to 20

for i  in range(1,21):
    print(f"Multiplication table of: {i}")
    for j in range(1,11):
        print(i*j)
        

# or a formatted table (nice and clean)

