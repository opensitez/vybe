<?php
// vybe-test: php/trait_conflict_resolution/as_changes_method_visibility_to_private
// origin: languages/php/tests/php/test_trait_conflict_resolution.rs
// vybe-test-mode: compile

trait Impl { public function helper(): string { return "h"; } }
class Widget {
    use Impl { helper as private internalHelper; }
    public function run(): string { return $this->internalHelper(); }
}
echo (new Widget())->run();
