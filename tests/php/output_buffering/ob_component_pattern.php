<?php
// vybe-test: php/output_buffering/ob_component_pattern
// origin: languages/php/tests/php/test_output_buffering.rs
// vybe-test-mode: compile

function component(callable $render): string {
    ob_start();
    $render();
    return ob_get_clean();
}
$output = component(function() {
    echo "Hello from component!";
});
echo $output;
