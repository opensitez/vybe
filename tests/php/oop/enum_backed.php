<?php
// vybe-test: php/oop/enum_backed
// origin: languages/php/tests/php/test_oop.rs
// vybe-test-mode: compile

enum Suit: string { case Hearts = 'H'; case Diamonds = 'D'; } echo Suit::Hearts->value;
