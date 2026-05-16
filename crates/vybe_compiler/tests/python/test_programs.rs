use super::helpers::{run_python, compile_ok};

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
fn word_frequency() {
    compile_ok(r#"
text = "the cat sat on the mat the cat"
words = text.split()
freq = {}
for w in words:
    if w in freq:
        freq[w] += 1
    else:
        freq[w] = 1
for k, v in freq.items():
    print(f"{k}: {v}")
"#);
}

#[test]
fn enumerate_pattern() {
    compile_ok(r#"
items = ["apple", "banana", "cherry"]
for i, item in enumerate(items):
    print(f"{i}: {item}")
"#);
}

#[test]
fn sorted_with_comp() {
    compile_ok(r#"
data = [5, 2, 8, 1, 9, 3]
ascending = sorted(data)
print(ascending)
total = sum(data)
avg = total / len(data)
print(f"sum={total}, avg={avg}")
"#);
}

#[test]
fn with_and_string_methods() {
    compile_ok(r#"
text = "  Hello, World!  "
cleaned = text.strip().lower()
words = cleaned.split()
print(len(words))
print("hello" in cleaned)
"#);
}

#[test]
fn nested_comprehension_real() {
    compile_ok(r#"
matrix = [[1,2,3],[4,5,6],[7,8,9]]
flat = [x for row in matrix for x in row]
evens = [x for x in flat if x % 2 == 0]
print(sorted(evens))
print(sum(evens))
"#);
}

// Runtime program tests

#[test]
fn fizzbuzz_runtime() {
    let out = run_python(r#"
for i in range(1, 16):
    if i % 15 == 0:
        print("FizzBuzz")
    elif i % 3 == 0:
        print("Fizz")
    elif i % 5 == 0:
        print("Buzz")
    else:
        print(i)
"#);
    assert_eq!(out, vec![
        "1", "2", "Fizz", "4", "Buzz", "Fizz", "7", "8", "Fizz", "Buzz",
        "11", "Fizz", "13", "14", "FizzBuzz"
    ]);
}

#[test]
fn fibonacci_runtime() {
    let out = run_python(r#"
def fib(n):
    if n <= 1:
        return n
    return fib(n - 1) + fib(n - 2)

for i in range(10):
    print(fib(i))
"#);
    assert_eq!(out, vec!["0", "1", "1", "2", "3", "5", "8", "13", "21", "34"]);
}

#[test]
fn factorial_runtime() {
    let out = run_python(r#"
def factorial(n):
    result = 1
    for i in range(1, n + 1):
        result *= i
    return result

print(factorial(10))
"#);
    assert_eq!(out, vec!["3628800"]);
}
