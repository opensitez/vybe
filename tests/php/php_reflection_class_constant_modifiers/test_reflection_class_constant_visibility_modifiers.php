<?php
// vybe-test: php/php_reflection_class_constant_modifiers/test_reflection_class_constant_visibility_modifiers
// origin: languages/php/tests/php/test_php_reflection_class_constant_modifiers.rs

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

class VisibilityDemo {
    public const PUB = 1;
    protected const PROT = 2;
    private const PRIV = 3;
}
$rc = new ReflectionClass(VisibilityDemo::class);
$pub = $rc->getReflectionConstant('PUB');
$prot = $rc->getReflectionConstant('PROT');
$priv = $rc->getReflectionConstant('PRIV');

echo ($pub->isPublic() ? '1' : '0') . ',' . ($prot->isProtected() ? '1' : '0') . ',' . ($priv->isPrivate() ? '1' : '0'), "\n";

__vybe_check(ob_get_clean(), "1,1,1");
