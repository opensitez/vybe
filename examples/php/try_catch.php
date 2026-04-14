<?php

function divide($a, $b) {
    if ($b == 0) {
        throw "Division by zero";
    }
    return $a / $b;
}

try {
    echo divide(10, 2);
    echo divide(10, 0);
} catch ($e) {
    echo "Error: " . $e;
} finally {
    echo "Done";
}
