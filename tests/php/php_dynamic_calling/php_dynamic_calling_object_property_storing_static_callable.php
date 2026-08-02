<?php
// vybe-test: php/php_dynamic_calling/php_dynamic_calling_object_property_storing_static_callable
// origin: languages/php/tests/php/test_php_dynamic_calling.rs

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

class StaticCarrier {
    public \Closure|string $dispatcher;
    public function __construct() {
        $this->dispatcher = [self::class, 'make'];
    }
    public static function make(int $n): int { return $n + 10; }
    public function run(int $n): int {
        $callable = $this->dispatcher;
        return $callable($n);
    }
}

$obj = new StaticCarrier();
echo $obj->run(5);

__vybe_check(ob_get_clean(), "15");
