<?php
// vybe-test: php/patterns/service_locator_resolves_by_name
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

class ServiceLocator {
    private static $services = [];
    public static function register(string $name, $service): void { self::$services[$name] = $service; }
    public static function get(string $name) { return self::$services[$name] ?? null; }
}
class Mailer {
    public function send(string $msg): void { echo 'mail:' . $msg; }
}
ServiceLocator::register('mailer', new Mailer());
$m = ServiceLocator::get('mailer');
$m->send('hello');

__vybe_check(ob_get_clean(), "mail:hello");
