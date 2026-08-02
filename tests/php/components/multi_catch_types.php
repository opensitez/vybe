<?php
// vybe-test: php/components/multi_catch_types
// origin: languages/php/tests/php/test_components.rs
// vybe-test-mode: compile

try {
    throw new Exception('oops');
} catch (TypeError | ValueError $e) {
    echo 'type or value error';
} catch (Exception $e) {
    echo 'generic';
}
