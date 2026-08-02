<?php
// vybe-test: php/catch_type_union_order/catch_interface_implemented_by_custom_exception
// origin: languages/php/tests/php/test_catch_type_union_order.rs

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

interface Loggable { public function logMessage(): string; }
class AuditEvent extends Exception implements Loggable {
    public function logMessage(): string { return 'audit'; }
}
try { throw new AuditEvent('evt'); }
catch (Loggable $l) { echo $l->logMessage(); }

__vybe_check(ob_get_clean(), "audit");
