<?php
// vybe-test: php/magic_methods/magic_get_chained
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

class Config {
    private array $data = ["db" => ["host" => "localhost"]];
    public function __get($key) {
        $val = $this->data[$key] ?? null;
        if (is_array($val)) {
            $child = new Config();
            foreach ($val as $k => $v) {
                $child->data[$k] = $v;
            }
            return $child;
        }
        return $val;
    }
    public function __set($key, $value) {
        $this->data[$key] = $value;
    }
}
$c = new Config();
echo $c->db->host;

__vybe_check(ob_get_clean(), "localhost");
