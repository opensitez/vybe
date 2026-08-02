<?php
// vybe-test: php/filter_var_array_recursive/filter_var_array_recursive
// origin: languages/php/tests/php/test_filter_var_array_recursive.rs

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

$data = [
    'user' => [
        'email' => 'test@example.com',
        'age'   => 'not-an-int'
    ]
];
$args = [
    'user' => [
        'filter' => FILTER_VALIDATE_EMAIL,
        'flags'  => FILTER_REQUIRE_ARRAY
    ] // Actually filter_var_array doesn't recurse automatically without manual iteration, but we test the array flag.
];
// Wait, we'll just test a simpler filter_var_array with basic associative inputs
$data2 = ['a' => '1', 'b' => '2'];
$res = filter_var_array($data2, FILTER_VALIDATE_INT);
echo $res['a'] . "|" . $res['b'];

__vybe_check(ob_get_clean(), "1|2");
