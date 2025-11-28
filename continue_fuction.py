# 5. Skip input check

# Ask the user to enter 5 values.
# If the user enters a negative number, skip it using continue.

# 6. Skip certain names

# Loop through a list:
# ["Ravi", "Amit", "Suresh", "Rohan"]
# Skip "Suresh" using continue.

name = ["Ravi", "Amit", "Suresh", "Rohan"]

for i in (name):
    if i =="Suresh":
        continue
    print(i)

# 7. Skip uppercase letters

# Given "PyThOnIsFuN"
# Print only lowercase letters.
# Use continue to skip uppercase ones.

character = "PyThOnIsFuN"
for i in character:
    if i <=character.lower():
        continue
    print(i)

# 8. Skip numbers greater than 50

# Loop from 1 to 100.
# Skip numbers greater than 50 using continue.

for i in range(1,101):
    if i>50:
        continue
    print(i)

# 9. Skip odd numbers except 1

# Loop through 1–20.
# Skip all odd numbers except 1.

# 10. Skip if divisible by both 2 and 5

# Loop through 1–50.
# Use continue when the number is divisible by both 2 and 5 (i.e., divisible by 10).

for i in  range(1,51):
    if i % 2 == 0 and i % 5 ==0:
        continue
    print(i)
    
    
#  ✅ Harder continue Question (Medium–Hard Level)
# 🔥 Question:

# You have a list of numbers:

# numbers = [12, 45, 0, -5, 23, 90, 0, 11, -2, 60]


# Write a loop that prints only the positive numbers BUT skip:

#  All zeros

# All numbers above 50

# All negative numbers

# You must use continue in your solution.

# 🧠 Expected Output:
# 12
# 45
# 23
# 11

numbers = [12, 45, 0, -5, 23, 90, 0, 11, -2, 60]

for i in numbers:
    if ((i == 0) or (i>=50) or (i<0)):
        continue
    print(i)