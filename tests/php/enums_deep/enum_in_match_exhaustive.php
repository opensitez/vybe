<?php
// vybe-test: php/enums_deep/enum_in_match_exhaustive
// origin: languages/php/tests/php/test_enums_deep.rs
// vybe-test-mode: compile

enum Direction { case North; case South; case East; case West; }
function opposite(Direction $d): Direction {
    return match($d) {
        Direction::North => Direction::South,
        Direction::South => Direction::North,
        Direction::East  => Direction::West,
        Direction::West  => Direction::East,
    };
}
echo opposite(Direction::North)->name;
echo opposite(Direction::East)->name;
