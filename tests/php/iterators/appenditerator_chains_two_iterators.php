<?php
// vybe-test: php/iterators/appenditerator_chains_two_iterators
// origin: languages/php/tests/php/test_iterators.rs

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

$app = new AppendIterator();
$app->append(new ArrayIterator([1, 2]));
$app->append(new ArrayIterator([3]));
echo implode('', iterator_to_array($app));

__vybe_check(ob_get_clean(), "32");
