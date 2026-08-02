<?php
// vybe-test: php/pcre_advanced/preg_match_all_offset_capture
// origin: languages/php/tests/php/test_pcre_advanced.rs
// vybe-test-mode: compile

preg_match_all('/\d+/', 'abc123def456', $m, PREG_OFFSET_CAPTURE);
echo $m[0][0][0] . '@' . $m[0][0][1];
echo $m[0][1][0] . '@' . $m[0][1][1];
