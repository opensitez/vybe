<?php
// vybe-test: php/oop_patterns/__call_forwarding_runtime
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

class Invoker {
    public function __call(string $name, array $args): string {
        return $name . ':' . implode(',', $args);
    }
}
$i = new Invoker();
echo $i->run('build', 1, 2);

__vybe_check(ob_get_clean(), "run:build,1,2");
