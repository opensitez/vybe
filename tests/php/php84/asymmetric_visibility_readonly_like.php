<?php
// vybe-test: php/php84/asymmetric_visibility_readonly_like
// origin: languages/php/tests/php/test_php84.rs
// vybe-test-mode: compile

class Config {
    public private(set) string $env = 'production';
    public private(set) bool $debug = false;
    public function setDev(): void { $this->env = 'dev'; $this->debug = true; }
}
$cfg = new Config();
echo $cfg->env . ':' . ($cfg->debug ? 'debug' : 'prod');
$cfg->setDev();
echo $cfg->env . ':' . ($cfg->debug ? 'debug' : 'prod');
