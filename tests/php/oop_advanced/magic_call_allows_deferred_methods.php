<?php
// vybe-test: php/oop_advanced/magic_call_allows_deferred_methods
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

class Broker {
    public function __call(string $name, array $args): string {
        return $name . ":" . implode(",", $args);
    }
}
$b = new Broker();
echo $b->enqueue("a", 1);
echo "|";
echo $b->flush();

__vybe_check(ob_get_clean(), "enqueue:a,1|flush:");
