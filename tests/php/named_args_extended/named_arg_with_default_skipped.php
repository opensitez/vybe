<?php
// vybe-test: php/named_args_extended/named_arg_with_default_skipped
// origin: languages/php/tests/php/test_named_args_extended.rs

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
        public string $host = 'localhost',
        public int $port = 3306,
        public string $dbname = 'default'
    ) {}
}
$c = new Config(dbname: 'myapp');
echo $c->host . ':' . $c->port . '/' . $c->dbname;

__vybe_check(ob_get_clean(), "localhost:3306/myapp");
