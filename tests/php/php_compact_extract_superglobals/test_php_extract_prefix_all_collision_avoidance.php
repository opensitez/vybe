<?php
// vybe-test: php/php_compact_extract_superglobals/test_php_extract_prefix_all_collision_avoidance
// origin: languages/php/tests/php/test_php_compact_extract_superglobals.rs

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

$name = "Original Name";
$params = ["name" => "New Name", "id" => 100];

extract($params, EXTR_PREFIX_ALL, "req");
echo "$name | $req_name | $req_id";

__vybe_check(ob_get_clean(), "Original Name | New Name | 100");
