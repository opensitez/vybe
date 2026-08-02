<?php
// vybe-test: php/trait_conflict_resolution/trait_with_constant_php82
// origin: languages/php/tests/php/test_trait_conflict_resolution.rs
// vybe-test-mode: compile

trait HasVersion {
    const VERSION = '1.0';
}
class App {
    use HasVersion;
}
echo App::VERSION;
