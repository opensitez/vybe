<?php
// vybe-test: php/serialization_advanced/sleep_selective_properties
// origin: languages/php/tests/php/test_serialization_advanced.rs

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

class DBConnection {
    private string $host;
    private int $port;
    private mixed $connection = null;
    public function __construct(string $host, int $port) {
        $this->host = $host; $this->port = $port;
    }
    public function __sleep(): array { return ['host', 'port']; }
    public function __wakeup(): void { $this->connection = null; }
    public function getHost(): string { return $this->host; }
}
$db = new DBConnection('localhost', 5432);
$s = serialize($db);
$db2 = unserialize($s);
echo $db2->getHost();

__vybe_check(ob_get_clean(), "localhost");
