<?php
// vybe-test: php/weak_references_runtime/weak_reference_singleton_factory
// origin: languages/php/tests/php/test_weak_references_runtime.rs

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

class Singleton {
    private static ?Singleton $instance = null;
    public static function getInstance(): static {
        if (static::$instance === null) {
            static::$instance = new static();
        }
        return static::$instance;
    }
    public static function createWeak(): WeakReference {
        return WeakReference::create(static::getInstance());
    }
}
$ref = Singleton::createWeak();
echo ($ref->get() instanceof Singleton) ? 'ok' : 'fail';

__vybe_check(ob_get_clean(), "ok");
