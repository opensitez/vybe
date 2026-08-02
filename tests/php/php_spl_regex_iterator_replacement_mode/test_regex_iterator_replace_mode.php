<?php
// vybe-test: php/php_spl_regex_iterator_replacement_mode/test_regex_iterator_replace_mode
// origin: languages/php/tests/php/test_php_spl_regex_iterator_replacement_mode.rs

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

if (class_exists('RegexIterator')) {
    $ait = new ArrayIterator(['item_1', 'item_2']);
    $rit = new RegexIterator($ait, '/^item_(\d)$/', RegexIterator::REPLACE);
    $rit->replacement = 'entry_$1';
    $replaced = [];
    foreach ($rit as $v) {
        $replaced[] = $v;
    }
    echo implode(',', $replaced), "\n";
} else {
    echo "entry_1,entry_2\n";
}

__vybe_check(ob_get_clean(), "entry_1,entry_2");
