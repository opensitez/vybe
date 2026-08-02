<?php
// vybe-test: php/oop/class_alias_and_exists_runtime
// origin: languages/php/tests/php/test_oop.rs

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

class OriginalWidget {
    public function label(): string { return 'orig'; }
}
class_alias('OriginalWidget', 'AliasWidget');
echo class_exists('AliasWidget') ? 'yes' : 'no';
echo '|';
echo (new AliasWidget())->label();

__vybe_check(ob_get_clean(), "yes|orig");
