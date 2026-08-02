<?php
// vybe-test: php/pcre_advanced/negative_lookahead
// origin: languages/php/tests/php/test_pcre_advanced.rs
// vybe-test-mode: compile

preg_match_all('/\bfoo(?!bar)\w*/', 'foobar foobaz fooqwe', $m);
echo implode(',', $m[0]);
