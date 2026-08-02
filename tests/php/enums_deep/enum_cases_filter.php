<?php
// vybe-test: php/enums_deep/enum_cases_filter
// origin: languages/php/tests/php/test_enums_deep.rs
// vybe-test-mode: compile

enum Priority: int {
    case Low    = 1;
    case Medium = 2;
    case High   = 3;
    case Critical = 4;
    public function isUrgent(): bool { return $this->value >= 3; }
}
$urgent = array_filter(Priority::cases(), fn($c) => $c->isUrgent());
echo count($urgent);
echo ':' . implode(',', array_map(fn($c) => $c->name, $urgent));
