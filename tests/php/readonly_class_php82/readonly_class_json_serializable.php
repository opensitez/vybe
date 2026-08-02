<?php
// vybe-test: php/readonly_class_php82/readonly_class_json_serializable
// origin: languages/php/tests/php/test_readonly_class_php82.rs

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

readonly class Product {
    public function __construct(
        public string $name,
        public float $price,
    ) {}
}
$p = new Product("Widget", 9.99);
echo json_encode(['name' => $p->name, 'price' => $p->price]);

__vybe_check(ob_get_clean(), "{\"name\":\"Widget\",\"price\":9.99}");
