<?php
// vybe-test: php/literals/test_php_array_literal_variants
// origin: languages/php/tests/php/test_literals.rs

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

$a = [1, 2, 3];
$b = ['x' => 10, 'y' => 20];
$c = [0 => 'zero', 2 => 'two', 'x' => 'ex'];

echo count($a);
echo '\n';
echo $b['x'];
echo '\n';
$merged = $a + [2 => 5, 3 => 6];
echo json_encode($merged);

__vybe_check(ob_get_clean(), "3\\n10\\n[1,2,3,6]");
