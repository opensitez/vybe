<?php
// vybe-test: php/oop/enum_basic
// origin: languages/php/tests/php/test_oop.rs
// vybe-test-mode: compile

enum Color { case Red; case Green; case Blue; } $c = Color::Red; echo $c->name;
