<?php
// vybe-test: php/oop_advanced/trait_with_static_property_runtime
// origin: languages/php/tests/php/test_oop_advanced.rs

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

trait Counter {
    protected static int $count = 0;
    public function next(): int {
        return ++self::$count;
    }
}
class A {
    use Counter;
}
class B {
    use Counter;
}
$a = new A();
$b = new B();
echo $a->next();
echo $b->next();

__vybe_check(ob_get_clean(), "11");
