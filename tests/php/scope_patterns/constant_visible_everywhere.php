<?php
// vybe-test: php/scope_patterns/constant_visible_everywhere
// origin: languages/php/tests/php/test_scope_patterns.rs
// vybe-test-mode: compile

define('SITE_NAME', 'Vybe');
function getSite(): string {
    return SITE_NAME;
}
class Config {
    public function name(): string { return SITE_NAME; }
}
echo getSite();
echo (new Config())->name();
