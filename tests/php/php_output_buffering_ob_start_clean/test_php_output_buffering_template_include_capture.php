<?php
// vybe-test: php/php_output_buffering_ob_start_clean/test_php_output_buffering_template_include_capture
// origin: languages/php/tests/php/test_php_output_buffering_ob_start_clean.rs
// vybe-test-mode: compile

function renderTemplate(string $content): string {
    ob_start();
    echo "<div>$content</div>";
    return ob_get_clean();
}

echo renderTemplate("Hello World");
