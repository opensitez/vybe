<?php
// vybe-test: php/enums_deep/enum_constants_basic
// origin: languages/php/tests/php/test_enums_deep.rs
// vybe-test-mode: compile

enum Suit: string {
    case Hearts   = 'H';
    case Diamonds = 'D';
    case Clubs    = 'C';
    case Spades   = 'S';
    const array RED_SUITS  = [self::Hearts, self::Diamonds];
    const array BLACK_SUITS = [self::Clubs, self::Spades];
}
echo count(Suit::RED_SUITS) . ':' . count(Suit::BLACK_SUITS);
