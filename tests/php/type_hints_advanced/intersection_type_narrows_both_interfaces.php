<?php
// vybe-test: php/type_hints_advanced/intersection_type_narrows_both_interfaces
// origin: languages/php/tests/php/test_type_hints_advanced.rs

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

interface Serializable2 { public function serialize2(): string; }
interface Loggable { public function log(): void; }
class Service implements Serializable2, Loggable {
    public function serialize2(): string { return 'data'; }
    public function log(): void { echo 'logged'; }
}
function process(Serializable2&Loggable $obj): string { $obj->log(); return $obj->serialize2(); }
echo process(new Service);

__vybe_check(ob_get_clean(), "loggeddata");
