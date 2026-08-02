# vybe-test: python/programs/fizzbuzz
# origin: languages/python/tests/python/test_programs.rs
# vybe-test-mode: compile

for i in range(1, 101):
    if i % 15 == 0:
        print("FizzBuzz")
    elif i % 3 == 0:
        print("Fizz")
    elif i % 5 == 0:
        print("Buzz")
    else:
        print(i)
