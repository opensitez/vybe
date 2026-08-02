<?php
// vybe-test: php/inheritance_runtime/trait_alias_changes_visibility
// origin: languages/php/tests/php/test_inheritance_runtime.rs

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

trait T { private function secret(): string { return 's'; } }
class C { use T { secret as public show; } }
echo (new C())->show();

__vybe_check(ob_get_clean(), "s");
