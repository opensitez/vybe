<?php
// vybe-test: php/references/return_by_reference_property
// origin: languages/php/tests/php/test_references.rs
// vybe-test-mode: compile

class Config {
    private array $data = [];
    public function &item(string $key): mixed {
        if (!isset($this->data[$key])) { $this->data[$key] = null; }
        return $this->data[$key];
    }
}
$cfg = new Config();
$ref = &$cfg->item('debug');
$ref = true;
echo $cfg->item('debug') ? 'on' : 'off';
