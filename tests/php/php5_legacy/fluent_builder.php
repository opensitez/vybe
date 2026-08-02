<?php
// vybe-test: php/php5_legacy/fluent_builder
// origin: languages/php/tests/php/test_php5_legacy.rs
// vybe-test-mode: compile

class Response {
    public $status = 200;
    public $body = '';
    public $headers = [];
    public static function create() { return new Response(); }
    public function status($code) { $this->status = $code; return $this; }
    public function body($content) { $this->body = $content; return $this; }
    public function header($key, $val) { array_push($this->headers, $key . ': ' . $val); return $this; }
}
$resp = Response::create()
    ->status(200)
    ->body('{"ok":true}')
    ->header('Content-Type', 'application/json');
echo $resp->body;
