<?php
// vybe-test: php/oop/object_to_array_and_cast_runtime
// origin: languages/php/tests/php/test_oop.rs

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

class Pair {
    public function __construct(public int $a, public int $b) {}
}
$p = new Pair(1, 2);
$arr = (array) $p;
ksort($arr);
echo json_encode($arr);

__vybe_check(ob_get_clean(), "{\"a\":1,\"b\":2}");
