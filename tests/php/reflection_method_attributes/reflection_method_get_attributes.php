<?php
// vybe-test: php/reflection_method_attributes/reflection_method_get_attributes
// origin: languages/php/tests/php/test_reflection_method_attributes.rs

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
class Route {
    public function __construct(public string $path) {}
}

class Controller {
    #[Route('/home')]
    public function index() {}
}

$rm = new ReflectionMethod(Controller::class, 'index');
$attrs = $rm->getAttributes();
echo $attrs[0]->getName() . "->";
echo $attrs[0]->getArguments()[0];

__vybe_check(ob_get_clean(), "Route->/home");
