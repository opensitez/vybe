<?php
// vybe-test: php/php84_property_hooks/property_hook_set_triggers_side_effect
// origin: languages/php/tests/php/test_php84_property_hooks.rs

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

class Observable {
    private array $listeners = [];
    public int $value {
        set(int $v) {
            $old = $this->value ?? null;
            $this->value = $v;
            foreach ($this->listeners as $fn) $fn($old, $v);
        }
    }
    public function onChange(callable $fn): void { $this->listeners[] = $fn; }
}
$o = new Observable();
$o->onChange(fn($old, $new) => print("changed: $new\n"));
$o->value = 10;
$o->value = 20;

__vybe_check(ob_get_clean(), "changed: 10\nchanged: 20");
