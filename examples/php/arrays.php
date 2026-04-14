<?php

// Indexed array
$numbers = [10, 20, 30, 40, 50];
echo $numbers[0];
echo $numbers[4];

// Associative array
$person = [
    "name" => "Alice",
    "age" => 30,
    "active" => true
];
echo $person["name"];
echo $person["age"];

// Modify array
$numbers[2] = 99;
echo $numbers[2];

$person["email"] = "alice@example.com";
echo $person["email"];

// Nested arrays
$matrix = [
    [1, 2, 3],
    [4, 5, 6],
    [7, 8, 9]
];
echo $matrix[0][0];
echo $matrix[1][1];
echo $matrix[2][2];

// Array as stack
$stack = [];
$stack[] = "first";
$stack[] = "second";
$stack[] = "third";
foreach ($stack as $item) {
    echo $item;
}

// Nested associative
$config = [
    "database" => [
        "host" => "localhost",
        "port" => 3306
    ],
    "app" => [
        "name" => "MyApp",
        "debug" => true
    ]
];
echo $config["database"]["host"];
echo $config["app"]["name"];
