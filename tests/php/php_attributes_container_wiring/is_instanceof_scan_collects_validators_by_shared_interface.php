<?php
// vybe-test: php/php_attributes_container_wiring/is_instanceof_scan_collects_validators_by_shared_interface
// origin: languages/php/tests/php/test_php_attributes_container_wiring.rs

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

interface Constraint {
    public function check(int $v): bool;
}
#[Attribute]
class Min implements Constraint {
    public function __construct(private int $n) {}
    public function check(int $v): bool { return $v >= $this->n; }
}
#[Attribute]
class Max implements Constraint {
    public function __construct(private int $n) {}
    public function check(int $v): bool { return $v <= $this->n; }
}
class Form {
    #[Min(5)]
    #[Max(10)]
    public int $age = 0;
}
$rp = new ReflectionProperty(Form::class, 'age');
$ok = [];
foreach ($rp->getAttributes(Constraint::class, ReflectionAttribute::IS_INSTANCEOF) as $a) {
    $ok[] = $a->newInstance()->check(7) ? 'y' : 'n';
}
echo count($ok) . ':' . implode('', $ok);

__vybe_check(ob_get_clean(), "2:yy");
