<?php
// vybe-test: php/type_functions_extended/ctype_lower_lowercase
// origin: languages/php/tests/php/test_type_functions_extended.rs
// vybe-test-mode: compile

echo ctype_lower('hello') ? 'yes' : 'no';
echo ctype_lower('Hello') ? 'yes' : 'no';
