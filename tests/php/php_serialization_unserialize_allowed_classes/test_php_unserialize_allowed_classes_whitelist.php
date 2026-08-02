<?php
// vybe-test: php/php_serialization_unserialize_allowed_classes/test_php_unserialize_allowed_classes_whitelist
// origin: languages/php/tests/php/test_php_serialization_unserialize_allowed_classes.rs

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

class AllowedDto {
    public string $name = "Alice";
}

class BlockedDto {
    public string $secret = "123";
}

$p1 = serialize(new AllowedDto());
$p2 = serialize(new BlockedDto());

$r1 = unserialize($p1, ["allowed_classes" => [AllowedDto::class]]);
$r2 = unserialize($p2, ["allowed_classes" => [AllowedDto::class]]);

echo get_class($r1) . " vs " . get_class($r2);

__vybe_check(ob_get_clean(), "AllowedDto vs __PHP_Incomplete_Class");
