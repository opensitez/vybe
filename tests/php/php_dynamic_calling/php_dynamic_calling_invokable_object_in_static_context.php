<?php
// vybe-test: php/php_dynamic_calling/php_dynamic_calling_invokable_object_in_static_context
// origin: languages/php/tests/php/test_php_dynamic_calling.rs

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

class Invokable {
    public function __invoke(string $name): string { return 'hi:' . $name; }
}
class Caller {
    public static function execute(callable $cb, string $label): string {
        return $cb($label);
    }
}
$obj = new Invokable();
echo Caller::execute($obj, 'php');

__vybe_check(ob_get_clean(), "hi:php");
