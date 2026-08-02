<?php
// vybe-test: php/oop_patterns/private_constructor_factory_runtime
// origin: languages/php/tests/php/test_oop_patterns.rs

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

class Database {
    private static ?self $instance = null;
    private function __construct(private string $dsn) {}
    public static function connect(string $dsn): self {
        if (self::$instance === null) {
            self::$instance = new self($dsn);
        }
        return self::$instance;
    }
    public function getDsn(): string { return $this->dsn; }
}
$db1 = Database::connect('mysql://localhost/app');
$db2 = Database::connect('other://ignore');
echo $db1->getDsn() . '|' . (($db1 === $db2) ? 'same' : 'diff');

__vybe_check(ob_get_clean(), "mysql://localhost/app|same");
