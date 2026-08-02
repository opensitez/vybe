<?php
// vybe-test: php/spl_autoload/autoload_call_count_when_class_loaded_once
// origin: languages/php/tests/php/test_spl_autoload.rs

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

$calls = 0;
spl_autoload_register(function (string $class) use (&$calls): void {
    if ($class === 'Cache\\Svc') {
        $calls++;
        eval('namespace Cache; class Svc { public function name(): string { return \"svc\"; } }');
    }
});
class_exists('Cache\\Svc');
class_exists('Cache\\Svc', false);
class_exists('Cache\\Svc');
echo $calls;

__vybe_check(ob_get_clean(), "1");
