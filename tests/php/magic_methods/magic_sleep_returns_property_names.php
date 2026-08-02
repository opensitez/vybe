<?php
// vybe-test: php/magic_methods/magic_sleep_returns_property_names
// origin: languages/php/tests/php/test_magic_methods.rs

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
    public string $dsn = "mysql:host=localhost";
    public string $status = "connected";
    public function __sleep(): array {
        return ["dsn"];
    }
    public function __wakeup(): void {
        $this->status = "reconnected";
    }
}
$c = new Connection();
$data = serialize($c);
$c2 = unserialize($data);
echo $c2->dsn;
echo $c2->status;

__vybe_check(ob_get_clean(), "mysql:host=localhostreconnected");
