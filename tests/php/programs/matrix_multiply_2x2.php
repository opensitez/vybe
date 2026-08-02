<?php
// vybe-test: php/programs/matrix_multiply_2x2
// origin: languages/php/tests/php/test_programs.rs

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

function matmul(array $a, array $b): array {
    $n = count($a);
    $res = array_fill(0, $n, array_fill(0, $n, 0));
    for ($i = 0; $i < $n; $i++)
        for ($j = 0; $j < $n; $j++)
            for ($k = 0; $k < $n; $k++)
                $res[$i][$j] += $a[$i][$k] * $b[$k][$j];
    return $res;
}
$a = [[1,2],[3,4]];
$b = [[5,6],[7,8]];
$c = matmul($a, $b);
echo $c[0][0] . ',' . $c[0][1] . "\n";
echo $c[1][0] . ',' . $c[1][1] . "\n";

__vybe_check(ob_get_clean(), "19,22\n43,50");
