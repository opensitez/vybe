<?php
// vybe-test: php/php_traits_composition_conflict_resolution/test_php_trait_composed_from_multiple_traits
// origin: languages/php/tests/php/test_php_traits_composition_conflict_resolution.rs

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

trait Timestampable {
    public string $createdAt = "2024-01-01";
}
trait SoftDeletable {
    public bool $isDeleted = false;
}
trait Auditable {
    use Timestampable, SoftDeletable;
}

class Article {
    use Auditable;
}

$a = new Article();
echo $a->createdAt . " deleted=" . ($a->isDeleted ? "1" : "0");

__vybe_check(ob_get_clean(), "2024-01-01 deleted=0");
