<?php
// vybe-test: php/php_anonymous_classes_runtime_usage/test_php_anonymous_class_nested_inside_class_method
// origin: languages/php/tests/php/test_php_anonymous_classes_runtime_usage.rs
// vybe-test-mode: compile

class Container {
    public function createStrategy(): object {
        return new class {
            public function execute(): string { return "Strategy executed"; }
        };
    }
}

$c = new Container();
$strat = $c->createStrategy();
echo $strat->execute();
