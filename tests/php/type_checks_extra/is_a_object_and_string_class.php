<?php
// vybe-test: php/type_checks_extra/is_a_object_and_string_class
// origin: languages/php/tests/php/test_type_checks_extra.rs
// vybe-test-mode: compile

class Cat {}
class Kitten extends Cat {}
$k = new Kitten();
echo is_a($k, 'Cat') ? 'yes' : 'no';
echo is_a($k, 'Kitten') ? 'yes' : 'no';
echo is_a('Kitten', 'Cat', true) ? 'yes' : 'no';
