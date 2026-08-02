<?php
// vybe-test: php/php_spl_infinite_iterator_cycling/test_php_spl_infinite_iterator_cycles_repeatedly
// origin: languages/php/tests/php/test_php_spl_infinite_iterator_cycling.rs

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

$arr = new ArrayIterator(["red", "blue"]);
$inf = new InfiniteIterator($arr);

$out = [];
$count = 0;
foreach ($inf as $val) {
    $out[] = $val;
    $count++;
    if ($count >= 5) break;
}
echo implode(",", $out);

__vybe_check(ob_get_clean(), "red,blue,red,blue,red");
