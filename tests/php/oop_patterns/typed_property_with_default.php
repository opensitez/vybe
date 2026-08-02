<?php
// vybe-test: php/oop_patterns/typed_property_with_default
// origin: languages/php/tests/php/test_oop_patterns.rs
// vybe-test-mode: compile

class Config {
    public string   $env      = 'production';
    public int      $maxRetry = 3;
    public bool     $debug    = false;
    public float    $timeout  = 30.0;
    public array    $tags     = [];
    public ?string  $secret   = null;
}
$c = new Config();
echo $c->env;
echo $c->maxRetry;
echo $c->debug ? 'true' : 'false';
echo $c->timeout;
echo $c->secret ?? 'null';
