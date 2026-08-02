<?php
// vybe-test: php/spl_autoload/spl_autoload_register_with_array_callback
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

class ArrayLoader {
    public static function load(string $class): void {
        if ($class === 'Array\\Svc') {
            eval('namespace Array; class Svc { public function name(): string { return \"array\"; } }');
        }
    }
}
spl_autoload_register([ArrayLoader::class, 'load']);
echo (new Array\Svc())->name();

__vybe_check(ob_get_clean(), "array");
