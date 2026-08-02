<?php
// vybe-test: php/match_expressions/match_arms_evaluate_until_first_match
// origin: languages/php/tests/php/test_match_expressions.rs

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

$log = [];
$first = function() use (&$log) { $log[] = 'first'; return 10; };
$second = function() use (&$log) { $log[] = 'second'; return 20; };
$third = function() use (&$log) { $log[] = 'third'; return 30; };
echo match (20) {
    $first() => '10',
    $second() => '20',
    $third() => '30',
    default => 'none',
};
echo '|';
echo implode(',', $log);

__vybe_check(ob_get_clean(), "20|first,second");
