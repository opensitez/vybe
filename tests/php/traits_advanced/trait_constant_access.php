<?php
// vybe-test: php/traits_advanced/trait_constant_access
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

trait HasVersion {
    public function getVersion(): string { return static::VERSION; }
}
class AppV1 {
    use HasVersion;
    const VERSION = '1.0.0';
}
echo (new AppV1)->getVersion();

__vybe_check(ob_get_clean(), "1.0.0");
