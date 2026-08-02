<?php
// vybe-test: php/match_advanced/match_in_loop
// origin: languages/php/tests/php/test_match_advanced.rs
// vybe-test-mode: compile

$words = ['hello', 'WORLD', 'PHP', 'code'];
$result = [];
foreach ($words as $w) {
    $result[] = match(true) {
        ctype_upper($w) => strtolower($w),
        ctype_lower($w) => strtoupper($w),
        default         => $w,
    };
}
echo implode(',', $result);
