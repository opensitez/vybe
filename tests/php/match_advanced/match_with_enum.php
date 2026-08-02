<?php
// vybe-test: php/match_advanced/match_with_enum
// origin: languages/php/tests/php/test_match_advanced.rs
// vybe-test-mode: compile

enum Suit { case Hearts; case Diamonds; case Clubs; case Spades; }
$suit = Suit::Hearts;
$color = match($suit) {
    Suit::Hearts, Suit::Diamonds => 'red',
    Suit::Clubs,  Suit::Spades   => 'black',
};
echo $color;
