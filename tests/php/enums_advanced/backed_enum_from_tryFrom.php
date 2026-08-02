<?php
// vybe-test: php/enums_advanced/backed_enum_from_tryFrom
// origin: languages/php/tests/php/test_enums_advanced.rs

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

enum Weekday: int { case Mon=1; case Tue=2; case Wed=3; case Thu=4; case Fri=5; case Sat=6; case Sun=7; }
echo Weekday::from(3)->name;
echo ',' . (Weekday::tryFrom(99)?->name ?? 'none');

__vybe_check(ob_get_clean(), "Wed,none");
