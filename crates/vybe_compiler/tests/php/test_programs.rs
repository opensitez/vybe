use super::helpers::compile_ok;

#[test]
fn fizzbuzz() {
    compile_ok(r#"<?php
for ($i = 1; $i <= 100; $i++) {
    if ($i % 15 == 0) echo "FizzBuzz\n";
    elseif ($i % 3 == 0) echo "Fizz\n";
    elseif ($i % 5 == 0) echo "Buzz\n";
    else echo $i . "\n";
}
"#);
}

#[test]
fn fibonacci() {
    compile_ok(r#"<?php
function fib($n) {
    if ($n <= 1) return $n;
    return fib($n - 1) + fib($n - 2);
}
for ($i = 0; $i < 10; $i++) {
    echo fib($i) . "\n";
}
"#);
}

#[test]
fn factorial() {
    compile_ok(r#"<?php
function factorial($n) {
    $result = 1;
    for ($i = 1; $i <= $n; $i++) {
        $result *= $i;
    }
    return $result;
}
echo factorial(10);
"#);
}

#[test]
fn bubble_sort() {
    compile_ok(r#"<?php
function bubbleSort($arr) {
    $n = count($arr);
    for ($i = 0; $i < $n; $i++) {
        for ($j = 0; $j < $n - $i - 1; $j++) {
            if ($arr[$j] > $arr[$j + 1]) {
                $tmp = $arr[$j];
                $arr[$j] = $arr[$j + 1];
                $arr[$j + 1] = $tmp;
            }
        }
    }
    return $arr;
}
$sorted = bubbleSort([5, 3, 8, 1, 2]);
"#);
}

#[test]
fn class_hierarchy() {
    compile_ok(r#"<?php
class Animal {
    public $name;
    public function __construct($name) {
        $this->name = $name;
    }
    public function speak() {
        return $this->name . ' makes a sound';
    }
}

class Dog extends Animal {
    public function speak() {
        return $this->name . ' says Woof';
    }
}

class Cat extends Animal {
    public function speak() {
        return $this->name . ' says Meow';
    }
}

$animals = [new Dog('Rex'), new Cat('Whiskers'), new Dog('Buddy')];
foreach ($animals as $animal) {
    echo $animal->speak() . "\n";
}
"#);
}

#[test]
fn closures_and_callbacks() {
    compile_ok(r#"<?php
$numbers = [1, 2, 3, 4, 5];
$doubled = array_map(fn($n) => $n * 2, $numbers);
$evens = array_filter($numbers, fn($n) => $n % 2 == 0);
$sum = array_reduce($numbers, fn($carry, $item) => $carry + $item, 0);
"#);
}

#[test]
fn string_processing() {
    compile_ok(r#"<?php
$name = "  John Doe  ";
$name = trim($name);
$parts = explode(' ', $name);
$first = $parts[0];
$last = $parts[1];
$upper = strtoupper($name);
$lower = strtolower($name);
$replaced = str_replace('John', 'Jane', $name);
$contains = str_contains($name, 'John');
$starts = str_starts_with($name, 'John');
$len = strlen($name);
echo implode(', ', [$first, $last, $upper, $lower]);
"#);
}

#[test]
fn array_operations() {
    compile_ok(r#"<?php
$fruits = ['apple', 'banana', 'cherry'];
array_push($fruits, 'date');
$last = array_pop($fruits);
$reversed = array_reverse($fruits);
$sliced = array_slice($fruits, 1, 2);
$found = in_array('banana', $fruits);
$idx = array_search('cherry', $fruits);
$merged = array_merge($fruits, ['elderberry', 'fig']);
$keys = array_keys(['a' => 1, 'b' => 2]);
$values = array_values(['a' => 1, 'b' => 2]);
"#);
}

#[test]
fn math_operations() {
    compile_ok(r#"<?php
$x = abs(-42);
$y = ceil(3.2);
$z = floor(3.8);
$w = round(3.5);
$p = pow(2, 10);
$s = sqrt(144);
$mx = max(1, 5, 3);
$mn = min(1, 5, 3);
$r = rand();
"#);
}

#[test]
fn switch_calculator() {
    compile_ok(r#"<?php
function calc($op, $a, $b) {
    switch ($op) {
        case '+': return $a + $b;
        case '-': return $a - $b;
        case '*': return $a * $b;
        case '/': return $a / $b;
        default: return null;
    }
}
echo calc('+', 10, 5);
echo calc('*', 3, 7);
"#);
}

#[test]
fn json_roundtrip() {
    compile_ok(r#"<?php
$data = ['name' => 'John', 'age' => 30, 'scores' => [95, 87, 92]];
$json = json_encode($data);
$decoded = json_decode($json);
"#);
}
