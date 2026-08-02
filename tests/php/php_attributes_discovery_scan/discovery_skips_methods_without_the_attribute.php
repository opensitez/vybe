<?php
// vybe-test: php/php_attributes_discovery_scan/discovery_skips_methods_without_the_attribute
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

#[Attribute]
class Route {
    public function __construct(public string $path) {}
}
class C {
    #[Route('/a')]
    public function a() {}
    public function b() {}
    #[Route('/c')]
    public function c() {}
}
$n = 0;
$paths = [];
foreach ((new ReflectionClass(C::class))->getMethods() as $m) {
    $n++;
    if ($a = $m->getAttributes(Route::class)) {
        $paths[] = $a[0]->newInstance()->path;
    }
}
echo $n . ':' . implode(',', $paths);

__vybe_check(ob_get_clean(), "3:/a,/c");
