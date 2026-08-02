<?php
// vybe-test: php/edge_cases/self_referencing
// origin: languages/php/tests/php/test_edge_cases.rs
// vybe-test-mode: compile

class Node { public $next; } $a = new Node(); $b = new Node(); $a->next = $b; $b->next = $a;
