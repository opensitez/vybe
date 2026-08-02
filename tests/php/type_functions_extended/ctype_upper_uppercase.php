<?php
// vybe-test: php/type_functions_extended/ctype_upper_uppercase
// origin: languages/php/tests/php/test_type_functions_extended.rs
// vybe-test-mode: compile

echo ctype_upper('HELLO') ? 'yes' : 'no';
echo ctype_upper('Hello') ? 'yes' : 'no';
