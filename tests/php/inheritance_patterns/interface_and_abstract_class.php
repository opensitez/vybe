<?php
// vybe-test: php/inheritance_patterns/interface_and_abstract_class
// origin: languages/php/tests/php/test_inheritance_patterns.rs

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

interface Cacheable { public function cacheKey(): string; }
abstract class BaseModel implements Cacheable {
    abstract public function id(): int;
    public function cacheKey(): string { return get_class($this) . ':' . $this->id(); }
}
class User4 extends BaseModel { public function id(): int { return 42; } }
echo (new User4)->cacheKey(), "\n";

__vybe_check(ob_get_clean(), "User4:42");
