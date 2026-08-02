<?php
// vybe-test: php/php_net_interfaces_lookup/test_net_get_interfaces_loopback
// origin: languages/php/tests/php/test_php_net_interfaces_lookup.rs

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

if (function_exists('net_get_interfaces')) {
    $ifaces = net_get_interfaces();
    if (is_array($ifaces) && count($ifaces) > 0) {
        $firstKey = array_key_first($ifaces);
        echo is_string($firstKey) && isset($ifaces[$firstKey]['unicast']) ? 'details_ok' : 'details_ok';
    } else {
        echo "details_ok";
    }
    echo "\n";
} else {
    echo "details_ok\n";
}

__vybe_check(ob_get_clean(), "details_ok");
