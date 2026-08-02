<?php
// vybe-test: php/php_oop_property_hooks_get_set_syntax/test_property_hooks_get_set_syntax
// origin: languages/php/tests/php/test_php_oop_property_hooks_get_set_syntax.rs

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

class UserHookDemo {
    private string $_first = 'John';
    private string $_last = 'Doe';

    public string $fullName {
        get => $this->_first . ' ' . $this->_last;
    }
}
$u = new UserHookDemo();
echo $u->fullName, "\n";

__vybe_check(ob_get_clean(), "John Doe");
