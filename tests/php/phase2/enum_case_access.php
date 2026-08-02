<?php
// vybe-test: php/phase2/enum_case_access
// origin: languages/php/tests/php/test_phase2.rs
// vybe-test-mode: compile

enum Color { case Red; case Green; case Blue; }
$c = Color::Red;
echo $c->name;
echo $c->value;
