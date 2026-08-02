<?php
// vybe-test: php/advanced_closures/closure_stored_as_event_handler
// origin: languages/php/tests/php/test_advanced_closures.rs
// vybe-test-mode: compile

class EventEmitter {
    private array $handlers = [];
    public function on(string $event, callable $fn): void { $this->handlers[$event] = $fn; }
    public function emit(string $event, mixed $data): void {
        if (isset($this->handlers[$event])) ($this->handlers[$event])($data);
    }
}
$emitter = new EventEmitter();
$log = [];
$emitter->on('data', function(mixed $d) use (&$log): void { $log[] = $d; });
$emitter->emit('data', 'payload');
echo count($log);
