<?php
// vybe-test: php/traits_advanced/trait_multiple_classes_share
// origin: languages/php/tests/php/test_traits_advanced.rs

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

trait Timestamped {
    private string $createdAt;
    public function setCreatedAt(string $t): void { $this->createdAt = $t; }
    public function getCreatedAt(): string { return $this->createdAt; }
}
class Post { use Timestamped; }
class Comment { use Timestamped; }
$p = new Post; $p->setCreatedAt('2024-01-01');
$c = new Comment; $c->setCreatedAt('2024-06-15');
echo $p->getCreatedAt() . ',' . $c->getCreatedAt();

__vybe_check(ob_get_clean(), "2024-01-01,2024-06-15");
