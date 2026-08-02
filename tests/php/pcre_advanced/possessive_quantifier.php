<?php
// vybe-test: php/pcre_advanced/possessive_quantifier
// origin: languages/php/tests/php/test_pcre_advanced.rs
// vybe-test-mode: compile

// PHP PCRE supports possessive via ++, *+, ?+
$pattern = '/^\w++$/';
echo preg_match($pattern, 'hello123') ? 'matched' : 'no match';
echo preg_match($pattern, 'hello world') ? 'matched' : 'no match';
