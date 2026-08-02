<?php
// vybe-test: php/enums_deep/enum_cases_map
// origin: languages/php/tests/php/test_enums_deep.rs
// vybe-test-mode: compile

enum Weekday: int {
    case Monday    = 1;
    case Tuesday   = 2;
    case Wednesday = 3;
    case Thursday  = 4;
    case Friday    = 5;
    case Saturday  = 6;
    case Sunday    = 7;
    public function isWeekend(): bool { return $this->value >= 6; }
}
$weekends = array_filter(Weekday::cases(), fn($d) => $d->isWeekend());
$names = array_map(fn($d) => $d->name, $weekends);
echo implode(',', $names);
