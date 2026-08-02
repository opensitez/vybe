<?php
// vybe-test: php/callables/closure_from_callable_on_magic_call
// origin: languages/php/tests/php/test_callables.rs

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

class Proxy {
    public function __call(string $name, array $args): string { return "proxy:$name"; }
}
$p = new Proxy();
try {
    $call = Closure::fromCallable([$p, 'anything']);
    echo $call();
} catch (Throwable $e) {
    echo 'err';
}

__vybe_check(ob_get_clean(), "proxy:anything");
