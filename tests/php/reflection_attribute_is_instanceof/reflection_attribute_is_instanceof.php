<?php
// vybe-test: php/reflection_attribute_is_instanceof/reflection_attribute_is_instanceof
// origin: languages/php/tests/php/test_reflection_attribute_is_instanceof.rs

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

interface BaseAttr {}

#[Attribute]
class ConcreteAttr implements BaseAttr {}

#[ConcreteAttr]
class Subject {}

$rc = new ReflectionClass(Subject::class);

// Filter by exact class
$attrs1 = $rc->getAttributes(ConcreteAttr::class);
echo count($attrs1) . "|";

// Filter by interface using IS_INSTANCEOF
$attrs2 = $rc->getAttributes(BaseAttr::class, ReflectionAttribute::IS_INSTANCEOF);
echo count($attrs2) . "|";

// Filter by interface without IS_INSTANCEOF (should be 0)
$attrs3 = $rc->getAttributes(BaseAttr::class);
echo count($attrs3);

__vybe_check(ob_get_clean(), "1|1|0");
