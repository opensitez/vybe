<?php
// vybe-test: php/match_advanced/match_on_array_element_and_missing_key_fallback
// origin: languages/php/tests/php/test_match_advanced.rs
// vybe-test-mode: compile

$payload = ['status' => 'ok'];
$label = match ($payload['status'] ?? 'unknown') {
    'ok' => 'good',
    'fail' => 'bad',
    default => 'unknown',
};
echo $label;
echo '|';
$label2 = match ($payload['retry'] ?? null) {
    null => 'no-retry',
    0 => 'zero',
    default => 'has',
};
echo $label2;
