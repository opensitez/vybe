<?php
// vybe-test: php/pcre_advanced/case_insensitive_with_unicode
// origin: languages/php/tests/php/test_pcre_advanced.rs
// vybe-test-mode: compile

echo preg_match('/héllo/iu', 'HÉLLO') ? 'matched' : 'no match';
echo preg_match('/\p{Lu}+/u', 'ABC') ? 'matched' : 'no match';
