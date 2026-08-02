<?php
// vybe-test: php/trait_conflict_resolution/as_changes_method_visibility_to_protected
// origin: languages/php/tests/php/test_trait_conflict_resolution.rs
// vybe-test-mode: compile

trait Impl { public function secret(): string { return "s"; } }
class Service {
    use Impl { secret as protected; }
}
