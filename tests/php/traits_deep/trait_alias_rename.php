<?php
// vybe-test: php/traits_deep/trait_alias_rename
// origin: languages/php/tests/php/test_traits_deep.rs
// vybe-test-mode: compile

trait Logger {
    public function log(string $msg): void { echo "[LOG] $msg"; }
}
class Service {
    use Logger { log as writeLog; }
    public function process(): void { $this->writeLog("processing"); }
}
(new Service())->process();
