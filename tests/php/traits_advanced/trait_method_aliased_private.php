<?php
// vybe-test: php/traits_advanced/trait_method_aliased_private
// origin: languages/php/tests/php/test_traits_advanced.rs

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

trait Helper { public function doWork(): string { return 'work'; } }
class Service {
    use Helper { doWork as private internalWork; }
    public function run(): string { return $this->internalWork(); }
}
echo (new Service)->run();

__vybe_check(ob_get_clean(), "work");
