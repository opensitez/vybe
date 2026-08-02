<?php
// vybe-test: php/php_object_chaining/php_chain_constructor_to_static_factory_runtime
// origin: languages/php/tests/php/test_php_object_chaining.rs

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

class Logger {
    public function __construct(public string $value) {}
    public static function from(string $value): static { return new static($value); }
    public function append(string $suffix): static { $this->value .= $suffix; return $this; }
}
$v = Logger::from('a')->append('b')->append('c');
echo $v->value;

__vybe_check(ob_get_clean(), "abc");
