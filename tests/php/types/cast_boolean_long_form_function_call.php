<?php
// vybe-test: php/types/cast_boolean_long_form_function_call
// origin: languages/php/tests/php/test_types.rs
// vybe-test-mode: compile

class Reader { public $currentTagContents = " 1 "; }
$reader = new Reader();
$value = (boolean)trim($reader->currentTagContents);
echo $value;
