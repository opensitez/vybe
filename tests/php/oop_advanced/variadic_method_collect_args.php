<?php
// vybe-test: php/oop_advanced/variadic_method_collect_args
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

class Formatter {
    public function format(string $tpl, mixed ...$args): string {
        return vsprintf($tpl, $args);
    }
}
$f = new Formatter();
echo $f->format("%s is %d years old", "Alice", 30), "\n";
echo $f->format("%.2f + %.2f = %.2f", 1.1, 2.2, 3.3), "\n";

__vybe_check(ob_get_clean(), "Alice is 30 years old\n1.10 + 2.20 = 3.30");
