<?php
// vybe-test: php/oop_advanced/union_type_property
// origin: languages/php/tests/php/test_oop_advanced.rs

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

class Response {
    public int|string $code;
    public function __construct(int|string $code) {
        $this->code = $code;
    }
    public function isOk(): bool {
        return $this->code === 200 || $this->code === "ok";
    }
}
$r1 = new Response(200);
$r2 = new Response("ok");
$r3 = new Response(500);
echo $r1->isOk() ? "ok" : "fail", "\n";
echo $r2->isOk() ? "ok" : "fail", "\n";
echo $r3->isOk() ? "ok" : "fail", "\n";

__vybe_check(ob_get_clean(), "ok\nok\nfail");
