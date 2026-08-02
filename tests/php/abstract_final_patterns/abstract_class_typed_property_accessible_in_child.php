<?php
// vybe-test: php/abstract_final_patterns/abstract_class_typed_property_accessible_in_child
// origin: languages/php/tests/php/test_abstract_final_patterns.rs

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

abstract class DataSource {
    protected string $connection = '';
    abstract public function connect(string $dsn): void;
    public function getConnection(): string { return $this->connection; }
}
class DbSource extends DataSource {
    public function connect(string $dsn): void { $this->connection = $dsn; }
}
$db = new DbSource();
$db->connect("mysql://localhost/mydb");
echo $db->getConnection(), "\n";

__vybe_check(ob_get_clean(), "mysql://localhost/mydb");
