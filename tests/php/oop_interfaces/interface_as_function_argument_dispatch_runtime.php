<?php
// vybe-test: php/oop_interfaces/interface_as_function_argument_dispatch_runtime
// origin: languages/php/tests/php/test_oop_interfaces.rs

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

interface Handler { public function handle(string $in): string; }
class Upper implements Handler { public function handle(string $in): string { return strtoupper($in); } }
class Lower implements Handler { public function handle(string $in): string { return strtolower($in); } }
function run_handler(Handler $h, string $in): string { return $h->handle($in); }
echo run_handler(new Upper(), 'abc') . '|' . run_handler(new Lower(), 'ABC');

__vybe_check(ob_get_clean(), "ABC|abc");
