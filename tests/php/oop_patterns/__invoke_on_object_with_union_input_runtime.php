<?php
// vybe-test: php/oop_patterns/__invoke_on_object_with_union_input_runtime
// origin: languages/php/tests/php/test_oop_patterns.rs

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

class Processor {
    public function __invoke(string|int $v): string {
        return (string)$v;
    }
}
$p = new Processor();
echo $p('id');
echo '|';
echo $p(12);

__vybe_check(ob_get_clean(), "id|12");
