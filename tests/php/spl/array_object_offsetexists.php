<?php
// vybe-test: php/spl/array_object_offsetexists
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$ao = new ArrayObject(['name' => 'Alice']);
echo $ao->offsetExists('name') ? 'yes' : 'no';
echo $ao->offsetExists('age')  ? 'yes' : 'no';
