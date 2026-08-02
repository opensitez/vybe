<?php
// vybe-test: php/pcre_advanced/preg_split_delim_capture
// origin: languages/php/tests/php/test_pcre_advanced.rs
// vybe-test-mode: compile

$parts = preg_split('/([\s,;]+)/', 'one, two; three four', -1, PREG_SPLIT_DELIM_CAPTURE);
// Result includes delimiters as captured groups
echo count($parts) > 4 ? 'has delimiters' : 'no delimiters';
echo $parts[0];
