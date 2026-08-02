<?php
// vybe-test: php/object_comparison/array_unique_with_equal_objects_keeps_first
// origin: languages/php/tests/php/test_object_comparison.rs

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

class Tag { public function __construct(public string $name) {} }
$tags = [new Tag('php'), new Tag('php'), new Tag('rust')];
$unique = array_unique($tags);
echo count($unique);

__vybe_check(ob_get_clean(), "2");
