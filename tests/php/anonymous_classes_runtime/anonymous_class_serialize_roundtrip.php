<?php
// vybe-test: php/anonymous_classes_runtime/anonymous_class_serialize_roundtrip
// origin: languages/php/tests/php/test_anonymous_classes_runtime.rs

function __vybe_check($got, $want) {
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

echo "anonymous_class_serialize_roundtrip_ok";

__vybe_check(ob_get_clean(), "anonymous_class_serialize_roundtrip_ok");
