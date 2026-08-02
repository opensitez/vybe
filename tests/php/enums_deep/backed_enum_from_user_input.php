<?php
// vybe-test: php/enums_deep/backed_enum_from_user_input
// origin: languages/php/tests/php/test_enums_deep.rs
// vybe-test-mode: compile

enum Language: string {
    case PHP    = 'php';
    case Python = 'python';
    case Rust   = 'rust';
}
$inputs = ['php', 'rust', 'javascript', 'python'];
foreach ($inputs as $input) {
    $lang = Language::tryFrom($input);
    echo ($lang !== null ? $lang->name : 'unknown') . ' ';
}
