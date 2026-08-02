<?php
// vybe-test: php/magic_methods/magic_get_lazy_property
// origin: languages/php/tests/php/test_magic_methods.rs

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

class LazyLoader {
    private array $computed = [];
    public function __get($name) {
        if (!isset($this->computed[$name])) {
            $this->computed[$name] = strtoupper($name) . "_VALUE";
        }
        return $this->computed[$name];
    }
}
$l = new LazyLoader();
echo $l->foo;
echo $l->bar;
echo $l->foo;

__vybe_check(ob_get_clean(), "FOO_VALUEBAR_VALUEFOO_VALUE");
