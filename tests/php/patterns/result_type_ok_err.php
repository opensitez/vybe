<?php
// vybe-test: php/patterns/result_type_ok_err
// origin: languages/php/tests/php/test_patterns.rs

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

class Result {
    private function __construct(private bool $ok, private $value, private string $error = '') {}
    public static function ok($value): self { return new self(true, $value); }
    public static function err(string $error): self { return new self(false, null, $error); }
    public function isOk(): bool { return $this->ok; }
    public function unwrap() { if (!$this->ok) throw new \Exception($this->error); return $this->value; }
    public function error(): string { return $this->error; }
}
function divide(int $a, int $b): Result {
    if ($b === 0) return Result::err('division by zero');
    return Result::ok($a / $b);
}
$r1 = divide(10, 2);
echo $r1->isOk() ? 'ok' : 'err';
echo $r1->unwrap();
$r2 = divide(5, 0);
echo $r2->isOk() ? 'ok' : 'err';
echo $r2->error();

__vybe_check(ob_get_clean(), "ok5errdivision by zero");
