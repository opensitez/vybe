<?php
// vybe-test: php/oop/oop_method_call_in_traits_with_override_runtime
// origin: languages/php/tests/php/test_oop.rs

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

trait Logger {
    public function tag(): string { return 'base'; }
}
trait PrefixLogger {
    public function stamp(string $value): string { return 'pref:' . $value; }
}
class Service {
    use Logger;
    public function value(): string { return $this->tag(); }
}
class ServiceWithPrefix extends Service {
    use PrefixLogger;
    public function value(): string { return $this->stamp(parent::value()); }
}
echo (new Service())->value();
echo '|';
echo (new ServiceWithPrefix())->value();

__vybe_check(ob_get_clean(), "base|pref:base");
