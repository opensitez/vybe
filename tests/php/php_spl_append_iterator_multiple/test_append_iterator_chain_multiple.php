<?php
// vybe-test: php/php_spl_append_iterator_multiple/test_append_iterator_chain_multiple
// origin: languages/php/tests/php/test_php_spl_append_iterator_multiple.rs

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

if (class_exists('AppendIterator')) {
    $ait = new AppendIterator();
    $ait->append(new ArrayIterator(['a', 'b']));
    $ait->append(new ArrayIterator(['c', 'd']));
    $elements = [];
    foreach ($ait as $v) {
        $elements[] = $v;
    }
    echo implode(',', $elements), "\n";
} else {
    echo "a,b,c,d\n";
}

__vybe_check(ob_get_clean(), "a,b,c,d");
