<?php
// vybe-test: php/closures_advanced/closure_bind_to_different_instance
// origin: languages/php/tests/php/test_closures_advanced.rs
// vybe-test-mode: compile

class Node { private string $label; public function __construct(string $l) { $this->label = $l; } }
$getLabel = function() { return $this->label; };
$a = new Node('alpha');
$b = new Node('beta');
$fa = $getLabel->bindTo($a, Node::class);
$fb = $getLabel->bindTo($b, Node::class);
echo $fa() . ',' . $fb();
