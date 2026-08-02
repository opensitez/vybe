<?php
// vybe-test: php/php_array_sorting_multisort_callbacks/test_php_usort_spaceship_operator
// origin: languages/php/tests/php/test_php_array_sorting_multisort_callbacks.rs

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

$users = [
    ["name" => "Bob", "age" => 30],
    ["name" => "Alice", "age" => 25],
    ["name" => "Charlie", "age" => 25],
];

usort($users, fn($a, $b) => $a["age"] <=> $b["age"] ?: $a["name"] <=> $b["name"]);
$sortedNames = array_column($users, "name");
echo implode(", ", $sortedNames);

__vybe_check(ob_get_clean(), "Alice, Charlie, Bob");
