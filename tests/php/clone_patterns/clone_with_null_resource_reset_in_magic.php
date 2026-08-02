<?php
// vybe-test: php/clone_patterns/clone_with_null_resource_reset_in_magic
// origin: languages/php/tests/php/test_clone_patterns.rs

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

class Connection {
    public ?string $handle = null;
    public function connect(): void { $this->handle = "connected"; }
    public function __clone() { $this->handle = null; }
}
$c = new Connection();
$c->connect();
$d = clone $c;
echo $c->handle . ',' . var_export($d->handle, true);

__vybe_check(ob_get_clean(), "connected,NULL");
