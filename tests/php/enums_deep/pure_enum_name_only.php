<?php
// vybe-test: php/enums_deep/pure_enum_name_only
// origin: languages/php/tests/php/test_enums_deep.rs
// vybe-test-mode: compile

enum Planet { case Mercury; case Venus; case Earth; case Mars; }
$names = array_map(fn($p) => $p->name, Planet::cases());
echo implode(',', $names);
