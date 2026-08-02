<?php
// vybe-test: php/php_resource_id_type_inspection/test_get_resource_id_stream
// origin: languages/php/tests/php/test_php_resource_id_type_inspection.rs

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

$f = fopen('php://memory', 'r+');
$id = get_resource_id($f);
fclose($f);
echo is_int($id) && $id > 0 ? 'id_ok' : 'err', "\n";

__vybe_check(ob_get_clean(), "id_ok");
