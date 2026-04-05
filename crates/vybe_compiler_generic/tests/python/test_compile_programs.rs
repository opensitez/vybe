use super::helpers::compile_ok;

#[test]
fn fizzbuzz() {
    compile_ok(r#"
for i in range(1, 101):
    if i % 15 == 0:
        print("FizzBuzz")
    elif i % 3 == 0:
        print("Fizz")
    elif i % 5 == 0:
        print("Buzz")
    else:
        print(i)
"#);
}

#[test]
fn fibonacci() {
    compile_ok(r#"
def fib(n):
    if n <= 1:
        return n
    return fib(n - 1) + fib(n - 2)

for i in range(10):
    print(fib(i))
"#);
}

#[test]
fn factorial() {
    compile_ok(r#"
def factorial(n):
    result = 1
    for i in range(1, n + 1):
        result *= i
    return result

print(factorial(10))
"#);
}

#[test]
fn bubble_sort() {
    compile_ok(r#"
def bubble_sort(lst):
    n = len(lst)
    for i in range(n):
        for j in range(0, n - i - 1):
            if lst[j] > lst[j + 1]:
                lst[j], lst[j + 1] = lst[j + 1], lst[j]
    return lst

print(bubble_sort([64, 34, 25, 12, 22, 11, 90]))
"#);
}

#[test]
fn counter_class() {
    compile_ok(r#"
class Counter:
    def __init__(self):
        self.count = 0

    def increment(self):
        self.count += 1

    def get(self):
        return self.count
"#);
}

#[test]
fn list_processing() {
    compile_ok(r#"
numbers = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
evens = [x for x in numbers if x % 2 == 0]
squares = [x ** 2 for x in evens]
total = 0
for s in squares:
    total += s
print(total)
"#);
}

#[test]
fn string_processing() {
    compile_ok(r#"
words = "hello world foo bar"
parts = words.split()
upper_parts = [w.upper() for w in parts]
result = " ".join(upper_parts)
print(result)
"#);
}

#[test]
fn try_except_real() {
    compile_ok(r#"
def safe_divide(a, b):
    try:
        return a / b
    except:
        print("division error")
        return 0

print(safe_divide(10, 2))
print(safe_divide(10, 0))
"#);
}

#[test]
fn nested_data_structures() {
    compile_ok(r#"
matrix = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
flat = [x for row in matrix for x in row]
print(flat)
"#);
}

#[test]
fn mixed_features() {
    compile_ok(r#"
import os

def process(items):
    results = []
    for item in items:
        if item > 0:
            results.append(item * 2)
    return results

data = [1, -2, 3, -4, 5]
output = process(data)
print(f"Processed: {output}")
"#);
}

#[test]
fn fstring_program() {
    compile_ok(r#"
name = "Alice"
age = 30
print(f"Name: {name}, Age: {age}")
print(f"Next year: {age + 1}")
"#);
}

#[test]
fn lambda_usage() {
    compile_ok(r#"
double = lambda x: x * 2
add = lambda x, y: x + y
print(double(5))
print(add(3, 4))
"#);
}
