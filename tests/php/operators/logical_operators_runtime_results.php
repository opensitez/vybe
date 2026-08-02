<?php
// vybe-test: php/operators/logical_operators_runtime_results
// origin: languages/php/tests/php/test_operators.rs

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

$_SERVER = [];
if (!isset($defaultLang) && !empty($_SERVER['HTTP_ACCEPT_LANGUAGE'])) {
	echo 'and-bad';
} else {
	echo 'and-ok';
}

if (!empty($_SERVER['HTTP_ACCEPT_LANGUAGE']) || !isset($defaultLang)) {
	echo 'or-ok';
} else {
	echo 'or-bad';
}

if (true and false) {
	echo 'word-and-bad';
} else {
	echo 'word-and-ok';
}

if (false or true) {
	echo 'word-or-ok';
} else {
	echo 'word-or-bad';
}

if (true xor true) {
	echo 'word-xor-bad';
} else {
	echo 'word-xor-ok';
}

__vybe_check(ob_get_clean(), "and-okor-okword-and-okword-or-okword-xor-ok");
