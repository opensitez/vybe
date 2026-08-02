<?php
// vybe-test: php/type_functions_extended/ctype_digit_all_digits
// origin: languages/php/tests/php/test_type_functions_extended.rs
// vybe-test-mode: compile

echo ctype_digit('12345') ? 'yes' : 'no';
echo ctype_digit('123a5') ? 'yes' : 'no';
