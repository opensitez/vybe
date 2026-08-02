<?php
// vybe-test: php/attributes/attribute_on_function_parameter
// origin: languages/php/tests/php/test_attributes.rs

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

#[Attribute]
class Sensitive {}
function login(string $user, #[Sensitive] string $pass): string {
    return $user;
}
$params = (new ReflectionFunction('login'))->getParameters();
echo count($params[1]->getAttributes(Sensitive::class));

__vybe_check(ob_get_clean(), "1");
