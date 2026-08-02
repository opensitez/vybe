<?php
// vybe-test: php/oop_advanced/trait_insteadof_with_alias
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

trait X {
    public function speak(): string { return "X speaks"; }
}
trait Y {
    public function speak(): string { return "Y speaks"; }
}
class Z {
    use X, Y {
        X::speak insteadof Y;
        Y::speak as ySpeak;
    }
}
$z = new Z();
echo $z->speak(), "\n";
echo $z->ySpeak(), "\n";

__vybe_check(ob_get_clean(), "X speaks\nY speaks");
