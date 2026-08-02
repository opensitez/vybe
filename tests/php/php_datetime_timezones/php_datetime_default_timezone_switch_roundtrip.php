<?php
// vybe-test: php/php_datetime_timezones/php_datetime_default_timezone_switch_roundtrip
// origin: languages/php/tests/php/test_php_datetime_timezones.rs

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

$before = date_default_timezone_get();
date_default_timezone_set('UTC');
$utc = date_default_timezone_get();
date_default_timezone_set('America/Los_Angeles');
$la = date_default_timezone_get();
date_default_timezone_set($before);
echo ($utc === 'UTC') ? 'utc' : 'not';
echo '|';
echo ($la === 'America/Los_Angeles') ? 'la' : 'no';

__vybe_check(ob_get_clean(), "utc|la");
