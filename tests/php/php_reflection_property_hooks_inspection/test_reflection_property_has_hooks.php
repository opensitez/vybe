<?php
// vybe-test: php/php_reflection_property_hooks_inspection/test_reflection_property_has_hooks
// origin: languages/php/tests/php/test_php_reflection_property_hooks_inspection.rs

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

class Example {
    public string $name;
}
$rp = new ReflectionProperty(Example::class, 'name');
if (method_exists($rp, 'hasHooks')) {
    echo $rp->hasHooks() ? 'has_hooks' : 'no_hooks', "\n";
} else {
    echo "no_hooks\n";
}

__vybe_check(ob_get_clean(), "no_hooks");
