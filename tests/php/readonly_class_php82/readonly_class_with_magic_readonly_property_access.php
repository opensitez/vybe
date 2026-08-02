<?php
// vybe-test: php/readonly_class_php82/readonly_class_with_magic_readonly_property_access
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

readonly class Profile {
    public function __construct(public string $name) {}
    public function __get(string $n): mixed {
        if ($n === 'label') return strtoupper($this->name);
        return null;
    }
}
$p = new Profile('alice');
echo $p->name;
echo '|';
echo $p->label;

__vybe_check(ob_get_clean(), "alice|ALICE");
