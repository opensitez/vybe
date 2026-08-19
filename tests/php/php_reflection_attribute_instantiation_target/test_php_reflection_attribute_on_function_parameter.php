<?php
// vybe-test: php/php_reflection_attribute_instantiation_target/test_php_reflection_attribute_on_function_parameter
// origin: languages/php/tests/php/test_php_reflection_attribute_instantiation_target.rs

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

#[Attribute(Attribute::TARGET_PARAMETER)]
class ValidateEmail {}

function registerUser(#[ValidateEmail] string $email) {}

$rp = new ReflectionParameter("registerUser", "email");
$attrs = $rp->getAttributes(ValidateEmail::class);
echo count($attrs);


__vybe_check(ob_get_clean(), "1");
