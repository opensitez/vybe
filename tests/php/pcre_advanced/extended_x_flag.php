<?php
// vybe-test: php/pcre_advanced/extended_x_flag
// origin: languages/php/tests/php/test_pcre_advanced.rs
// vybe-test-mode: compile

$pattern = '/
    ^           # start
    \d{4}       # year
    -           # separator
    \d{2}       # month
    -           # separator
    \d{2}       # day
    $           # end
/x';
echo preg_match($pattern, '2024-06-15') ? 'matched' : 'no match';
echo preg_match($pattern, '24-6-1') ? 'matched' : 'no match';
