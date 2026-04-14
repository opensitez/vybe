<?php

$day = 3;

switch ($day) {
    case 1:
        echo "Monday";
        break;
    case 2:
        echo "Tuesday";
        break;
    case 3:
        echo "Wednesday";
        break;
    case 4:
        echo "Thursday";
        break;
    case 5:
        echo "Friday";
        break;
    default:
        echo "Weekend";
        break;
}

// Fall-through
$grade = "B";
switch ($grade) {
    case "A":
    case "B":
        echo "Excellent";
        break;
    case "C":
        echo "Average";
        break;
    default:
        echo "Below average";
        break;
}
