<?php

$score = 85;

if ($score >= 90) {
    echo "Grade: A";
} elseif ($score >= 80) {
    echo "Grade: B";
} elseif ($score >= 70) {
    echo "Grade: C";
} else {
    echo "Grade: F";
}

// Comparison operators
$a = 5;
$b = "5";
echo $a == $b;   // true (loose)
echo $a === $b;  // false (strict)
echo $a != $b;   // false
echo $a !== $b;  // true

// Logical operators
$x = true;
$y = false;
echo $x && $y;
echo $x || $y;
echo !$x;

// Ternary
$age = 20;
$label = ($age >= 18) ? "adult" : "minor";
echo $label;

// Null coalescing
$config = null;
$value = $config ?? "default";
echo $value;

// Match expression
$status = 200;
$text = match($status) {
    200 => "OK",
    404 => "Not Found",
    500 => "Server Error",
    default => "Unknown"
};
echo $text;
