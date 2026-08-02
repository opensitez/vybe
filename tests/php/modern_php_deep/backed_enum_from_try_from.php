<?php
// vybe-test: php/modern_php_deep/backed_enum_from_try_from
// origin: languages/php/tests/php/test_modern_php_deep.rs

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

enum Status: string {
    case Active   = "active";
    case Inactive = "inactive";
    case Banned   = "banned";
}
$s = Status::from("active");
echo $s->name;
$t = Status::tryFrom("unknown");
echo $t === null ? "null" : $t->name;

__vybe_check(ob_get_clean(), "Activenull");
