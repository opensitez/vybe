<?php
// vybe-test: php/intersection_types/intersection_interface_static_not_part_of_intersection
// origin: languages/php/tests/php/test_intersection_types.rs

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

interface Activatable { public function activate(): void; }
interface Deactivatable { public function deactivate(): void; }
class Toggle implements Activatable, Deactivatable {
    private bool $on = false;
    public function activate(): void { $this->on = true; }
    public function deactivate(): void { $this->on = false; }
    public function isOn(): bool { return $this->on; }
}
function toggle(Activatable&Deactivatable $t): void {
    $t->activate();
    $t->deactivate();
}
$t = new Toggle();
toggle($t);
echo $t->isOn() ? 'on' : 'off';

__vybe_check(ob_get_clean(), "off");
