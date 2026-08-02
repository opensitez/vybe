<?php
// vybe-test: php/filters/validate_form_input
// origin: languages/php/tests/php/test_filters.rs
// vybe-test-mode: compile

function validateUser(array $input): array {
    $errors = [];
    if (filter_var($input['email'] ?? '', FILTER_VALIDATE_EMAIL) === false) {
        $errors[] = 'Invalid email';
    }
    if (filter_var($input['age'] ?? '', FILTER_VALIDATE_INT,
        ['options' => ['min_range' => 0, 'max_range' => 150]]) === false) {
        $errors[] = 'Invalid age';
    }
    return $errors;
}
$valid   = validateUser(['email' => 'bob@example.com', 'age' => '25']);
$invalid = validateUser(['email' => 'notanemail', 'age' => '200']);
echo count($valid) . ':' . count($invalid);
