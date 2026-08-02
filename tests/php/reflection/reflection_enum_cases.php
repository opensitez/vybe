<?php
// vybe-test: php/reflection/reflection_enum_cases
// origin: languages/php/tests/php/test_reflection.rs
// vybe-test-mode: compile

enum Color { case Red; case Green; case Blue; }
$re = new ReflectionEnum(Color::class);
$cases = $re->getCases();
echo count($cases);
echo ':' . $cases[0]->getName();
