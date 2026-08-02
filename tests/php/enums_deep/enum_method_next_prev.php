<?php
// vybe-test: php/enums_deep/enum_method_next_prev
// origin: languages/php/tests/php/test_enums_deep.rs
// vybe-test-mode: compile

enum Month: int {
    case January  = 1; case February = 2; case March    = 3;
    case April    = 4; case May      = 5; case June     = 6;
    case July     = 7; case August   = 8; case September = 9;
    case October  = 10; case November = 11; case December = 12;
    public function next(): self {
        $next = ($this->value % 12) + 1;
        return self::from($next);
    }
    public function daysInMonth(int $year = 2024): int {
        return cal_days_in_month(CAL_GREGORIAN, $this->value, $year);
    }
}
echo Month::December->next()->name;
echo Month::February->daysInMonth(2024);
