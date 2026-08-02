<?php
// vybe-test: php/php_attributes_discovery_scan/class_level_prefix_combines_with_method_routes_across_classes
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

#[Attribute(Attribute::TARGET_CLASS)]
class Prefix {
    public function __construct(public string $base) {}
}
#[Attribute(Attribute::TARGET_METHOD)]
class Route {
    public function __construct(public string $path) {}
}
#[Prefix('/api/users')]
class UserApi {
    #[Route('/list')]
    public function list() {}
}
#[Prefix('/api/posts')]
class PostApi {
    #[Route('/list')]
    public function list() {}
    #[Route('/new')]
    public function new_() {}
}
$out = [];
foreach ([UserApi::class, PostApi::class] as $c) {
    $rc = new ReflectionClass($c);
    $pre = $rc->getAttributes(Prefix::class)[0]->newInstance()->base;
    foreach ($rc->getMethods() as $m) {
        foreach ($m->getAttributes(Route::class) as $a) {
            $out[] = $pre . $a->newInstance()->path;
        }
    }
}
echo implode(' ', $out);

__vybe_check(ob_get_clean(), "/api/users/list /api/posts/list /api/posts/new");
