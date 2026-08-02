<?php
// vybe-test: php/date_advanced/age_calculation
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

function calculateAge(DateTimeImmutable $birthday, DateTimeImmutable $today): int {
    return $birthday->diff($today)->y;
}
$birthday = new DateTimeImmutable('1990-06-15');
$today    = new DateTimeImmutable('2024-06-15');
echo calculateAge($birthday, $today);
