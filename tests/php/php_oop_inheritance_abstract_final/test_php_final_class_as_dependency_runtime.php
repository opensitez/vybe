<?php
// vybe-test: php/php_oop_inheritance_abstract_final/test_php_final_class_as_dependency_runtime
// origin: languages/php/tests/php/test_php_oop_inheritance_abstract_final.rs
// vybe-test-mode: compile

final class FinalService {
    public function execute(): string { return 'ok'; }
}
class Wrapped {
    public function call(FinalService $s): string { return $s->execute(); }
}
$w = new Wrapped();
echo $w->call(new FinalService());
