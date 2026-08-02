<?php
// vybe-test: php/datetime/test_timezone_constructor_and_offset_change
// origin: languages/php/tests/php/test_datetime.rs

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

$tz = new DateTimeZone('Europe/Paris');
$offset = $tz->getOffset(new DateTime('2024-07-01', $tz));
echo $offset > 3000 ? 'summer' : 'winter';
echo '|';
echo $tz->getName();

__vybe_check(ob_get_clean(), "summer|Europe/Paris");
