<?php
// vybe-test: php/callables/arrow_function_this_in_instance_method
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

class Host {
    private string $tag = 'inner';
    public function run(): string {
        $f = fn(): string => $this->tag;
        return $f();
    }
}
echo (new Host())->run();

__vybe_check(ob_get_clean(), "inner");
