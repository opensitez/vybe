<?php
// vybe-test: php/modern_php_deep/enum_in_type_hint
// origin: languages/php/tests/php/test_modern_php_deep.rs
// vybe-test-mode: compile

enum Season { case Spring; case Summer; case Autumn; case Winter; }
function describe(Season $s): string {
    return match($s) {
        Season::Spring => "flowers",
        Season::Summer => "sun",
        Season::Autumn => "leaves",
        Season::Winter => "snow" };
}
echo describe(Season::Winter);
