<?php

// While loop
$i = 0;
while ($i < 5) {
    echo $i;
    $i++;
}

// Do-while
$j = 0;
do {
    echo $j;
    $j++;
} while ($j < 3);

// For loop
for ($k = 10; $k > 0; $k -= 2) {
    echo $k;
}

// Foreach over array
$fruits = ["apple", "banana", "cherry"];
foreach ($fruits as $fruit) {
    echo $fruit;
}

// Foreach with key => value
$colors = ["red" => "#FF0000", "green" => "#00FF00", "blue" => "#0000FF"];
foreach ($colors as $name => $hex) {
    echo $name . ": " . $hex;
}

// Break and continue
for ($n = 0; $n < 10; $n++) {
    if ($n % 2 == 0) {
        continue;
    }
    if ($n > 7) {
        break;
    }
    echo $n;
}
