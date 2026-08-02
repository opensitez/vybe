<?php
// vybe-test: php/spl_autoload/invokable_object_loader_receives_class_name
// origin: languages/php/tests/php/test_spl_autoload.rs

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

class Collector {
    public array $seen = [];
    public function __invoke(string $class): void {
        $this->seen[] = $class;
        if ($class === 'AutoLoad\\Loaded') {
            eval('namespace AutoLoad; class Loaded {}');
        }
    }
}
$loader = new Collector();
spl_autoload_register($loader);
class_exists('AutoLoad\\Loaded');
echo $loader->seen[0] === 'AutoLoad\\Loaded' ? 'ok' : 'bad';

__vybe_check(ob_get_clean(), "ok");
