<?php
// vybe-test: php/oop_patterns/class_const_visibility_and_inheritance_runtime
// origin: languages/php/tests/php/test_oop_patterns.rs

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

class BaseValue {
    public const SCOPE = 'public';
    protected const INTERNAL = 'internal';
    private const SECRET = 'secret';
    public function marker(): string {
        return self::SCOPE . '|' . static::SCOPE;
    }
}
class ChildValue extends BaseValue {}
echo ChildValue::SCOPE;
echo '|' . (new ChildValue())->marker();

__vybe_check(ob_get_clean(), "public|public|public");
