<?php
// vybe-test: php/patterns/multiton_named_singletons
// origin: languages/php/tests/php/test_patterns.rs

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
    private static $pool = [];
    private function __construct(public string $name) {}
    public static function getInstance(string $name): self {
        if (!isset(self::$pool[$name])) {
            self::$pool[$name] = new self($name);
        }
        return self::$pool[$name];
    }
}
$a = Connection::getInstance('primary');
$b = Connection::getInstance('primary');
$c = Connection::getInstance('replica');
echo ($a === $b) ? 'same' : 'diff';
echo ($a === $c) ? 'same' : 'diff';
echo $c->name;

__vybe_check(ob_get_clean(), "samediffreplica");
