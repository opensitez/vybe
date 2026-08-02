<?php
// vybe-test: php/phase2/enum_backed_access
// origin: languages/php/tests/php/test_phase2.rs
// vybe-test-mode: compile

enum Suit: string {
    case Hearts = 'H';
    case Diamonds = 'D';
    case Clubs = 'C';
    case Spades = 'S';
}
$s = Suit::Hearts;
echo $s->value;
echo $s->name;
