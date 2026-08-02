<?php
// vybe-test: php/php_reflection_named_type_inspection/test_reflection_named_type_is_builtin
// origin: languages/php/tests/php/test_php_reflection_named_type_inspection.rs

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

class CustomObj {}
function handle(CustomObj $obj, string $str): void {}
$rf = new ReflectionFunction('handle');
$p1 = $rf->getParameters()[0]->getType();
$p2 = $rf->getParameters()[1]->getType();
echo ($p1->isBuiltin() ? 'p1_builtin' : 'p1_class') . ',' . ($p2->isBuiltin() ? 'p2_builtin' : 'p2_class'), "\n";

__vybe_check(ob_get_clean(), "p1_class,p2_builtin");
