<?php
// vybe-test: php/type_functions_extended/ctype_punct_punctuation
// origin: languages/php/tests/php/test_type_functions_extended.rs
// vybe-test-mode: compile

echo ctype_punct('!@#') ? 'yes' : 'no';
echo ctype_punct('!a#') ? 'yes' : 'no';
