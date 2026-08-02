<?php
// vybe-test: php/php_is_a_string_class_allow_string/test_is_a_string_class_name_allow_string_false
// origin: languages/php/tests/php/test_php_is_a_string_class_allow_string.rs

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

class ParentType2 {}
class ChildType2 extends ParentType2 {}
echo is_a('ChildType2', 'ParentType2', false) ? 'unexpected' : 'string_disallowed', "\n";

__vybe_check(ob_get_clean(), "string_disallowed");
