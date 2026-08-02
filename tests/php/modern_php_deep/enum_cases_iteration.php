<?php
// vybe-test: php/modern_php_deep/enum_cases_iteration
// origin: languages/php/tests/php/test_modern_php_deep.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

enum Direction {
    case North;
    case South;
    case East;
    case West;
}
$cases = Direction::cases();
echo count($cases);
echo $cases[0]->name;
echo $cases[3]->name;

__vybe_check(ob_get_clean(), "4NorthWest");
