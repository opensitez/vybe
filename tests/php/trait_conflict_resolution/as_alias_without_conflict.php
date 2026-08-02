<?php
// vybe-test: php/trait_conflict_resolution/as_alias_without_conflict
// origin: languages/php/tests/php/test_trait_conflict_resolution.rs

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

trait Logger { public function log(string $msg): void { echo $msg; } }
class App {
    use Logger { log as writeLog; }
}
(new App())->writeLog("hello");

__vybe_check(ob_get_clean(), "hello");
