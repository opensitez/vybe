<?php
// vybe-test: php/classes/class_trait_with_method_alias_runtime
// origin: languages/php/tests/php/test_classes.rs

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

trait Logger {
    public function message(): string { return 'base'; }
}
class Service {
    use Logger {
        message as public aliasMessage;
    }
}
echo (new Service())->aliasMessage();

__vybe_check(ob_get_clean(), "base");
