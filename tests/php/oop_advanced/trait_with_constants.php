<?php
// vybe-test: php/oop_advanced/trait_with_constants
// origin: languages/php/tests/php/test_oop_advanced.rs

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
    const VERSION = "1.0";
    public function getVersion(): string {
        return self::VERSION;
    }
}
class App {
    use HasVersion;
}
$app = new App();
echo $app->getVersion(), "\n";

__vybe_check(ob_get_clean(), "1.0");
