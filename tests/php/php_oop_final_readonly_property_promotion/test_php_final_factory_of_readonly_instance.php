<?php
// vybe-test: php/php_oop_final_readonly_property_promotion/test_php_final_factory_of_readonly_instance
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

final class Builder {
    private function __construct(
        public readonly string $token,
        public readonly int $ttl
    ) {}

    public static function fromDate(string $prefix, int $year): self {
        return new self("$prefix-$year", 900);
    }
}

$builder = Builder::fromDate("token", 2026);
echo $builder->token . "|" . $builder->ttl;

__vybe_check(ob_get_clean(), "token-2026|900");
