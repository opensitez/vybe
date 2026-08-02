<?php
// vybe-test: php/intersection_types/runtime_interface_check_simulates_intersection
// origin: languages/php/tests/php/test_intersection_types.rs

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

interface Loggable2 { public function log(): string; }
interface Auditable { public function audit(): string; }
class Event implements Loggable2, Auditable {
    public function log(): string { return "log"; }
    public function audit(): string { return "audit"; }
}
function verify(object $obj): string {
    if (!($obj instanceof Loggable2 && $obj instanceof Auditable)) return 'invalid';
    return $obj->log() . '+' . $obj->audit();
}
echo verify(new Event());

__vybe_check(ob_get_clean(), "log+audit");
