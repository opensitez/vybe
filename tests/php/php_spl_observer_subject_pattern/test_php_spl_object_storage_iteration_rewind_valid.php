<?php
// vybe-test: php/php_spl_observer_subject_pattern/test_php_spl_object_storage_iteration_rewind_valid
// origin: languages/php/tests/php/test_php_spl_observer_subject_pattern.rs
// vybe-test-mode: compile

$s = new SplObjectStorage();
$s->attach(new stdClass(), 1);
$s->attach(new stdClass(), 2);

$s->rewind();
$count = 0;
while ($s->valid()) {
    $count++;
    $s->next();
}
echo "Iterated $count items";
