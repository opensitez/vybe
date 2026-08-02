<?php
// vybe-test: php/php_traits_composition_conflict_resolution/test_php_trait_basic_method_inclusion
// origin: languages/php/tests/php/test_php_traits_composition_conflict_resolution.rs

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

trait Loggable {
    public function log(string $msg): string {
        return "[LOG] $msg";
    }
}

class User {
    use Loggable;
}

$u = new User();
echo $u->log("User created");

__vybe_check(ob_get_clean(), "[LOG] User created");
