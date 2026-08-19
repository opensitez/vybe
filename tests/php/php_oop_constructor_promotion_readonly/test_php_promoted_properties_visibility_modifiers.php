<?php
// vybe-test: php/php_oop_constructor_promotion_readonly/test_php_promoted_properties_visibility_modifiers
// origin: languages/php/tests/php/test_php_oop_constructor_promotion_readonly.rs

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

class Service {
    public function __construct(
        private string $secretKey,
        protected string $endpoint,
        public int $timeout = 30
    ) {}
    
    public function getEndpoint(): string {
        return $this->endpoint;
    }
}

$s = new Service("key_123", "https://api.example.com");
echo $s->getEndpoint();


__vybe_check(ob_get_clean(), "https://api.example.com");
