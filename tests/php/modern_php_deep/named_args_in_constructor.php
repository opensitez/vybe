<?php
// vybe-test: php/modern_php_deep/named_args_in_constructor
// origin: languages/php/tests/php/test_modern_php_deep.rs

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

class Config {
    public function __construct(
        public string $host = "localhost",
        public int    $port = 3306,
        public string $db   = "app"
    ) {}
}
$c = new Config(port: 5432, db: "mydb");
echo $c->host;
echo $c->port;
echo $c->db;

__vybe_check(ob_get_clean(), "localhost5432mydb");
