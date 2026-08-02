<?php
// vybe-test: php/php_serialize_allowed_classes_array/test_unserialize_allowed_classes_false
// origin: languages/php/tests/php/test_php_serialize_allowed_classes_array.rs

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

class PayloadDto {
    public int $id = 1;
}
$s = serialize(new PayloadDto());
$obj = unserialize($s, ['allowed_classes' => false]);
echo ($obj instanceof __PHP_Incomplete_Class) ? 'disallowed_all_ok' : 'err', "\n";

__vybe_check(ob_get_clean(), "disallowed_all_ok");
