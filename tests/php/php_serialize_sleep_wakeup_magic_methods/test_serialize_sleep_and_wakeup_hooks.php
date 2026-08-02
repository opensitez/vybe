<?php
// vybe-test: php/php_serialize_sleep_wakeup_magic_methods/test_serialize_sleep_and_wakeup_hooks
// origin: languages/php/tests/php/test_php_serialize_sleep_wakeup_magic_methods.rs

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

class SerializableObj {
    public string $keep = 'saved';
    public string $discard = 'ignored';
    public bool $restored = false;

    public function __sleep(): array {
        return ['keep'];
    }

    public function __wakeup(): void {
        $this->restored = true;
    }
}

$s = serialize(new SerializableObj());
$obj = unserialize($s);
echo $obj->keep . '|' . ($obj->restored ? 'restored' : 'not_restored'), "\n";

__vybe_check(ob_get_clean(), "saved|restored");
