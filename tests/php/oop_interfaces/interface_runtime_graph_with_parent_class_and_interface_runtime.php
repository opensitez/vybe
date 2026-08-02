<?php
// vybe-test: php/oop_interfaces/interface_runtime_graph_with_parent_class_and_interface_runtime
// origin: languages/php/tests/php/test_oop_interfaces.rs

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

interface A { public function a(): string; }
interface B extends A { public function b(): string; }
class ParentSvc {
    public function p(): string { return 'p'; }
}
class ChildSvc extends ParentSvc implements B {
    public function a(): string { return 'a'; }
    public function b(): string { return 'b'; }
}
$o = new ChildSvc();
echo $o->a() . $o->b() . $o->p();
echo '|';
echo ($o instanceof B) ? 'ok' : 'bad';

__vybe_check(ob_get_clean(), "abp|ok");
