<?php
// vybe-test: php/patterns/singleton_same_instance
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

class Config {
    private static $instance = null;
    public $data = [];
    private function __construct() {}
    public static function getInstance() {
        if (Config::$instance === null) {
            Config::$instance = new Config();
        }
        return Config::$instance;
    }
    public function set($k, $v) { $this->data[$k] = $v; }
    public function get($k) { return $this->data[$k] ?? null; }
}
$a = Config::getInstance();
$a->set('env', 'prod');
$b = Config::getInstance();
echo $b->get('env');
echo ($a === $b) ? 'same' : 'different';

__vybe_check(ob_get_clean(), "prodsame");
