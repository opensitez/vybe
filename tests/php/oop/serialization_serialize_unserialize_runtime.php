<?php
// vybe-test: php/oop/serialization_serialize_unserialize_runtime
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

class Package {
    public function __construct(public string $name, public int $version) {}
}
$p = new Package('core', 1);
$copy = unserialize(serialize($p));
echo $copy->name . '|' . $copy->version;

__vybe_check(ob_get_clean(), "core|1");
