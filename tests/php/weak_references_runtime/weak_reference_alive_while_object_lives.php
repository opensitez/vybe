<?php
// vybe-test: php/weak_references_runtime/weak_reference_alive_while_object_lives
// origin: languages/php/tests/php/test_weak_references_runtime.rs

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

class Resource { public string $name; public function __construct(string $n) { $this->name = $n; } }
$res = new Resource('db_conn');
$weak = WeakReference::create($res);
$alive = $weak->get();
echo ($alive !== null ? 'alive' : 'collected') . ':' . $weak->get()->name;

__vybe_check(ob_get_clean(), "alive:db_conn");
