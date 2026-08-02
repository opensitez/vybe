<?php
// vybe-test: php/union_types/union_type_in_property
// origin: languages/php/tests/php/test_union_types.rs

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

class Container { public int|string $value; }
$c = new Container; $c->value = 'text';
echo $c->value;
$c->value = 42;
echo ',' . $c->value;

__vybe_check(ob_get_clean(), "text,42");
