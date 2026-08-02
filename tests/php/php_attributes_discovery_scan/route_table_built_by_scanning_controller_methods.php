<?php
// vybe-test: php/php_attributes_discovery_scan/route_table_built_by_scanning_controller_methods
// origin: languages/php/tests/php/test_php_attributes_discovery_scan.rs

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

#[Attribute(Attribute::TARGET_METHOD)]
class Route {
    public function __construct(public string $path, public string $method = 'GET') {}
}
class UserController {
    #[Route('/users')]
    public function index(): string { return 'index'; }
    #[Route('/users/{id}')]
    public function show(): string { return 'show'; }
    public function helper(): string { return 'helper'; }
    #[Route('/users', method: 'POST')]
    public function create(): string { return 'create'; }
}
$routes = [];
foreach ((new ReflectionClass(UserController::class))->getMethods() as $m) {
    foreach ($m->getAttributes(Route::class) as $a) {
        $r = $a->newInstance();
        $routes[] = $r->method . ' ' . $r->path . ' -> ' . $m->getName();
    }
}
echo implode('|', $routes);

__vybe_check(ob_get_clean(), "GET /users -> index|GET /users/{id} -> show|POST /users -> create");
