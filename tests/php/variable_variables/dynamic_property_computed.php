<?php
// vybe-test: php/variable_variables/dynamic_property_computed
// origin: languages/php/tests/php/test_variable_variables.rs
// vybe-test-mode: compile

class Record {
    public string $field1 = 'a';
    public string $field2 = 'b';
    public string $field3 = 'c';
}
$r = new Record();
$result = '';
for ($i = 1; $i <= 3; $i++) {
    $result .= $r->{"field$i"};
}
echo $result;
