<?php
// vybe-test: php/array_offset_access/foreach_on_array_with_numeric_string_keys
// origin: languages/php/tests/php/test_array_offset_access.rs

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

$x = ['0' => 'zero', 1 => 'one', '2' => 'two'];
$out = [];
foreach ($x as $v) { $out[] = $v; }
echo implode('|', $out);

__vybe_check(ob_get_clean(), "zero|one|two");
