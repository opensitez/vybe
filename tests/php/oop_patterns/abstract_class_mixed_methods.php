<?php
// vybe-test: php/oop_patterns/abstract_class_mixed_methods
// origin: languages/php/tests/php/test_oop_patterns.rs
// vybe-test-mode: compile

abstract class Template {
    abstract protected function step1(): string;
    abstract protected function step2(): string;
    public function run(): string {
        return $this->step1() . '|' . $this->step2();
    }
    public function describe(): string { return 'Template'; }
}
class ConcreteTemplate extends Template {
    protected function step1(): string { return 'init'; }
    protected function step2(): string { return 'execute'; }
}
$t = new ConcreteTemplate();
echo $t->run();
echo $t->describe();
