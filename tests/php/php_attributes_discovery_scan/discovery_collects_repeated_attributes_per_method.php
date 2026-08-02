<?php
// vybe-test: php/php_attributes_discovery_scan/discovery_collects_repeated_attributes_per_method
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

#[Attribute(Attribute::IS_REPEATABLE | Attribute::TARGET_METHOD)]
class Verb {
    public function __construct(public string $name) {}
}
class Api {
    #[Verb('GET')]
    #[Verb('HEAD')]
    public function read(): string { return 'r'; }
}
$verbs = [];
foreach ((new ReflectionClass(Api::class))->getMethods() as $m) {
    foreach ($m->getAttributes(Verb::class) as $a) {
        $verbs[] = $a->newInstance()->name;
    }
}
echo implode(',', $verbs);

__vybe_check(ob_get_clean(), "GET,HEAD");
