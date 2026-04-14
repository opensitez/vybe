<?php

// Basic function
function greet($name) {
    return "Hello, " . $name . "!";
}
echo greet("World");

// Default parameters
function power($base, $exp = 2) {
    $result = 1;
    for ($i = 0; $i < $exp; $i++) {
        $result *= $base;
    }
    return $result;
}
echo power(3);
echo power(2, 10);

// Recursive function
function factorial($n) {
    if ($n <= 1) {
        return 1;
    }
    return $n * factorial($n - 1);
}
echo factorial(6);

// Closure
$double = function($x) {
    return $x * 2;
};
echo $double(21);

// Closure with use
$multiplier = 3;
$multiply = function($x) use ($multiplier) {
    return $x * $multiplier;
};
echo $multiply(7);

// Arrow function
$square = fn($x) => $x * $x;
echo $square(9);

// Higher-order function
function apply($fn, $val) {
    return $fn($val);
}
echo apply($square, 5);
echo apply(fn($x) => $x + 1, 99);

// Fibonacci
function fibonacci($n) {
    if ($n <= 1) {
        return $n;
    }
    return fibonacci($n - 1) + fibonacci($n - 2);
}
for ($i = 0; $i < 10; $i++) {
    echo fibonacci($i);
}
