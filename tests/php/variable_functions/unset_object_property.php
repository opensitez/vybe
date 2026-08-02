<?php
// vybe-test: php/variable_functions/unset_object_property
// origin: languages/php/tests/php/test_variable_functions.rs
// vybe-test-mode: compile

class Bag {
    public string $item = 'apple';
    public int    $qty  = 5;
}
$b = new Bag();
echo isset($b->item) ? 'set' : 'gone';
unset($b->item);
echo isset($b->item) ? 'set' : 'gone';
echo $b->qty;
