<?php
// vybe-test: php/spl_autoload/autoload_function_name_lookup_with_string_callable
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

function spl_autoload_string_loader(string $class): void {
    if ($class === 'Fn\\Service') {
        eval('namespace Fn; class Service { public function value(): string { return \"fn\"; } }');
    }
}
spl_autoload_register('spl_autoload_string_loader');
echo (new Fn\Service())->value();

__vybe_check(ob_get_clean(), "fn");
