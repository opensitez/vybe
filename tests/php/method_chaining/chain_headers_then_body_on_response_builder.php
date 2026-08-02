<?php
// vybe-test: php/method_chaining/chain_headers_then_body_on_response_builder
// origin: languages/php/tests/php/test_method_chaining.rs

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
    private array $headers = [];
    private string $body = '';
    public function header(string $k, string $v): static {
        $this->headers[] = "$k:$v";
        return $this;
    }
    public function body(string $b): static { $this->body = $b; return $this; }
    public function summary(): string {
        return count($this->headers) . '|' . $this->body;
    }
}
echo (new Response())->header('X', '1')->header('Y', '2')->body('ok')->summary();

__vybe_check(ob_get_clean(), "2|ok");
