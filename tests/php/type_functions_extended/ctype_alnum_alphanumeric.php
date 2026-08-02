<?php
// vybe-test: php/type_functions_extended/ctype_alnum_alphanumeric
// origin: languages/php/tests/php/test_type_functions_extended.rs
// vybe-test-mode: compile

echo ctype_alnum('abc123') ? 'yes' : 'no';
echo ctype_alnum('abc!23') ? 'yes' : 'no';
