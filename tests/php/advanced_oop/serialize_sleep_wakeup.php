<?php
// vybe-test: php/advanced_oop/serialize_sleep_wakeup
// origin: languages/php/tests/php/test_advanced_oop.rs

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

class Conn {
    public string $host = 'localhost';
    private mixed $resource = null;
    public function __sleep(): array { return ['host']; }
    public function __wakeup(): void { $this->resource = 'reconnected'; }
    public function status(): string { return $this->resource ?? 'null'; }
}
$c = new Conn;
$c2 = unserialize(serialize($c));
echo $c2->host . ':' . $c2->status();

__vybe_check(ob_get_clean(), "localhost:reconnected");
