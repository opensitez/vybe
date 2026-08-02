<?php
// vybe-test: php/php_spl_observer_subject_pattern/test_php_spl_observer_multiple_listeners
// origin: languages/php/tests/php/test_php_spl_observer_subject_pattern.rs
// vybe-test-mode: compile

class EventBus implements SplSubject {
    private SplObjectStorage $obs;
    public function __construct() { $this->obs = new SplObjectStorage(); }
    public function attach(SplObserver $o): void { $this->obs->attach($o); }
    public function detach(SplObserver $o): void { $this->obs->detach($o); }
    public function notify(): void { foreach ($this->obs as $o) { $o->update($this); } }
}

class LoggerObs implements SplObserver { public function update(SplSubject $s): void {} }
class MetricsObs implements SplObserver { public function update(SplSubject $s): void {} }

$bus = new EventBus();
$bus->attach(new LoggerObs());
$bus->attach(new MetricsObs());
$bus->notify();
