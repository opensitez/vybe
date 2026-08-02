<?php
// vybe-test: php/scope_variables/global_function_variable_and_static_property
// origin: languages/php/tests/php/test_scope_variables.rs

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

$callback = function (): string {
    global $tag;
    $tag = 'from-scope';
    return $tag;
};
$callback();
echo $callback();
echo '|';
class Holder {
    public static $tag = null;
}
Holder::$tag = 'global-tag';
function read_holder(): string { return Holder::$tag; }
echo read_holder();

__vybe_check(ob_get_clean(), "from-scope|from-scope|global-tag");
