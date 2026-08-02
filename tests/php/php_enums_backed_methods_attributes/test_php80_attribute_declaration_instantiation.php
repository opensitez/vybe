<?php
// vybe-test: php/php_enums_backed_methods_attributes/test_php80_attribute_declaration_instantiation
// origin: languages/php/tests/php/test_php_enums_backed_methods_attributes.rs

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

#[Attribute(Attribute::TARGET_CLASS | Attribute::TARGET_METHOD)]
class Route {
    public function __construct(
        public string $path,
        public string $method = "GET"
    ) {}
}

#[Route("/api/users", method: "POST")]
class UserController {}

$rc = new ReflectionClass(UserController::class);
$attrs = $rc->getAttributes(Route::class);
$route = $attrs[0]->newInstance();

echo "{$route->method} {$route->path}";

__vybe_check(ob_get_clean(), "POST /api/users");
