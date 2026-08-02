<?php
// vybe-test: php/php_oop_property_hooks_get_set_syntax/test_property_hooks_block_setter_runtime
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

class Logger {
    private string $_prefix = '';

    public string $tag {
        get => $this->_prefix;
        set {
            $this->_prefix = strtoupper($value);
        }
    }
}
$x = new Logger();
$x->tag = 'dev';
echo $x->tag;

__vybe_check(ob_get_clean(), "DEV");
