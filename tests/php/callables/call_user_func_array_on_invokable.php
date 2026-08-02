<?php
// vybe-test: php/callables/call_user_func_array_on_invokable
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

class Handler {
    public function __invoke(int $n): string { return 'value:' . $n; }
}
$h = new Handler();
echo call_user_func_array($h, [7]);

__vybe_check(ob_get_clean(), "value:7");
