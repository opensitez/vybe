<?php
// vybe-test: php/php_oop_final_readonly_property_promotion/test_php_promoted_property_doc_comments
// origin: languages/php/tests/php/test_php_oop_final_readonly_property_promotion.rs

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

class Customer {
    public function __construct(
        /** @var string Customer full name */
        public string $name,
        /** @var string Customer email address */
        public string $email
    ) {}
}

$c = new Customer("Bob", "bob@example.com");
echo "$c->name <$c->email>";


__vybe_check(ob_get_clean(), "Bob <bob@example.com>");
