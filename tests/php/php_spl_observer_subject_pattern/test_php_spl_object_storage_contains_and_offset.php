<?php
// vybe-test: php/php_spl_observer_subject_pattern/test_php_spl_object_storage_contains_and_offset
// origin: languages/php/tests/php/test_php_spl_observer_subject_pattern.rs

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

$storage = new SplObjectStorage();
$o1 = new stdClass();
$o2 = new stdClass();

$storage->attach($o1, "data1");
$storage->attach($o2, "data2");

echo ($storage->contains($o1) ? "1" : "0") . " | data=" . $storage[$o1];

__vybe_check(ob_get_clean(), "1 | data=data1");
