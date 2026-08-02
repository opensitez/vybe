<?php
// vybe-test: php/magic_method_errors/magic_isset_false_for_missing
// origin: languages/php/tests/php/test_magic_method_errors.rs

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

class Dyn {
    public function __isset(string $k): bool { return false; }
}
$d = new Dyn();
echo isset($d->ghost) ? 'yes' : 'no';

__vybe_check(ob_get_clean(), "no");
