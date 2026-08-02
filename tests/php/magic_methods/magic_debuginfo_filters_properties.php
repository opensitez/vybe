<?php
// vybe-test: php/magic_methods/magic_debuginfo_filters_properties
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

class Secret {
    public string $name = "visible";
    private string $password = "hidden";
    public function __debugInfo(): array {
        return ["name" => $this->name, "password" => "***"];
    }
}
$s = new Secret();
$info = $s->__debugInfo();
echo $info["name"];
echo $info["password"];

__vybe_check(ob_get_clean(), "visible***");
