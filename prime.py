# # # # # # # def get_average(numbers):
# # # # # # #     if len(numbers)==0:
# # # # # # #         return 0
    
# # # # # # #     total = sum(numbers)
# # # # # # #     avg = total/len(numbers)
# # # # # # #     return avg

# # # # # # # value = [1,23,4]
# # # # # # # print(get_average(value))



# # # # # # # # Question 1: Create a function to find the maximum of three numbers.

# # # # # # # # Example:
# # # # # # # # Input → 10, 5, 20
# # # # # # # # Output → 20

# # # # # # def max_numbers(numbers):
# # # # # #     return max(numbers)

# # # # # # value = [1,423,23,9]
# # # # # # print(max_numbers(value))


# # # # # # Question 2: Create a function that checks if a number is even or odd.

# # # # # # Return "Even" or "Odd".


# # # # # def even_or_odd(number):
# # # # #     if number % 2 == 0:
# # # # #         return ("even")
# # # # #     else:
# # # # #         return ("add num")
        
# # # # # user = int(input("enter a number to chek or or add: "))

# # # # # print(even_or_odd(user)) 

# # # # Question 3: Create a function that takes a list and returns only the even numbers.

# # # # Example:
# # # # Input → [1,2,3,4,5,6]
# # # # Output → [2,4,6]

# # def even(numbers):
# #     list = []
    
# #     for i in numbers:
# #         if i%2 == 0:
# #             l1 = list.append(i)
# #     return list  
# # values = [1,3,2,6]
# # print(even(values))        
         
    
    
# # # Question 4: Create a function that counts vowels in a string.

# # # Example:
# # # Input → "Ravi"
# # # Output → 2


# # def vowels(num):
# #     for i in num:
    
# #     return i   

# # name = "ravi"
# # print(vowels(name))

# # # 📌 Question 5: Create a function that reverses a string.

# # # # Input → "Python"
# # # # Output → "nohtyP"

# # def reverse(integers):
    
# #     reversed_text = integers[::-1]
# #     print(reversed_text)
    
# # name = "Ravi"
# # print(reverse(name))
    

# # # 📌 Question 6: Create a function that returns the square of every number in a list.

# # # Input → [1,2,3]
# # # Output → [1,4,9]

# # def square(numbers):
    
# #     list = []
    
# #     for i in numbers:
# #         sq = int(i) * int(i)
# #         list.append(sq)
# #     return list

# # num = [2,4,6]

# # print(square(num))
        

# # # 📌 Question 7: Create a function to calculate the average of numbers in a list.

# # # (You just learned this 🙂)


# # def avg(numbers):
    
# #     avg_num = sum(numbers)/len(numbers)
# #     print(round(avg_num,0))
    
# # df = [1,34,5]
# # print(avg(df))
    


# # # 📌 Question 8: Create a function that returns factorial of a number.

# # # Example:
# # # 5 → 120

# # import math
# # print(math.factorial(5))

# def fact(num):
#     fact1 = 1
    
#     while num>0:
#         fact1 = num * fact1
#         print(f"{fact1} x {num} = {fact1*num}")
#         num -=1
        
#     return fact1 

# num = 12  
# print(fact(num))

# # a = 5
# # print(factorial(a))    

# # # **📌 Question 9: Create a function that accepts a name and prints:

# # # “Hello, <name>! Welcome.”**

# # # 📌 Question 10: Create a function that returns the length of the longest word in a sentence.

def log(word):
    word1 = word.split()
    
    for i in word1:
        if len(i) == max(len(i)):
            print(f"{i}")
name = "I am Ravi"
print(log(name))

# # # Example: