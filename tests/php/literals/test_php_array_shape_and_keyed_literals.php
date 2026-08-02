<?php
// vybe-test: php/literals/test_php_array_shape_and_keyed_literals
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

$a = [0 => 'zero', 'x' => 1, 2 => 'two'];
$b = [
    'left' => ['k' => 1],
    'right' => ['v' => 2],
];
array_push($a, 'tail');
echo $a[0];
echo '\n';
echo $a['x'];
echo '\n';
echo $b['left']['k'];
echo '\n';
echo $a[3];

__vybe_check(ob_get_clean(), "zero\n1\n1\ntail");
