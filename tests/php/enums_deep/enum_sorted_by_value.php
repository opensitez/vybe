<?php
// vybe-test: php/enums_deep/enum_sorted_by_value
// origin: languages/php/tests/php/test_enums_deep.rs
// vybe-test-mode: compile

enum Priority: int { case Low = 1; case Medium = 5; case High = 10; }
$tasks = [
    ['name' => 'cleanup', 'priority' => Priority::Low],
    ['name' => 'deploy',  'priority' => Priority::High],
    ['name' => 'review',  'priority' => Priority::Medium],
];
usort($tasks, fn($a, $b) => $b['priority']->value <=> $a['priority']->value);
foreach ($tasks as $task) { echo $task['name'] . ' '; }
