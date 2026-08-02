<?php
// vybe-test: php/php_array_sorting_multisort_callbacks/test_php_usort_by_priority_then_name
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

$tasks = [
    ["name" => "db", "priority" => 2],
    ["name" => "api", "priority" => 2],
    ["name" => "jobs", "priority" => 1],
];

usort($tasks, function($a, $b) {
    if ($a["priority"] === $b["priority"]) {
        return strcmp($a["name"], $b["name"]);
    }
    return $a["priority"] <=> $b["priority"];
});

echo $tasks[0]["name"] . "|" . $tasks[1]["name"] . "|" . $tasks[2]["name"];

__vybe_check(ob_get_clean(), "jobs|api|db");
