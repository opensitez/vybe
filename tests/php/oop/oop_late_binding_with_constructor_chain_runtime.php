<?php
// vybe-test: php/oop/oop_late_binding_with_constructor_chain_runtime
// origin: languages/php/tests/php/test_oop.rs

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

class Root {
    public function __construct(public string $kind) {}
    public static function make(string $kind): static {
        return new static($kind);
    }
}
class Leaf extends Root {}
echo (new Leaf('leaf'))->kind;
echo '|';
echo Leaf::make('mk')->kind;

__vybe_check(ob_get_clean(), "leaf|mk");
