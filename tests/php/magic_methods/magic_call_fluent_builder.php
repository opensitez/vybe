<?php
// vybe-test: php/magic_methods/magic_call_fluent_builder
// origin: languages/php/tests/php/test_magic_methods.rs

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

class Query {
    private array $parts = [];
    public function __call($name, $args) {
        $this->parts[] = "$name:" . implode(",", $args);
        return $this;
    }
    public function build() {
        return implode(" | ", $this->parts);
    }
}
$q = new Query();
echo $q->select("id", "name")->from("users")->where("active=1")->build();

__vybe_check(ob_get_clean(), "select:id,name | from:users | where:active=1");
