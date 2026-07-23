use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: SPL Observer & Subject Design Pattern — SplObserver, SplSubject, SplObjectStorage attachment, notify()
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_spl_observer_subject_event_notification() {
    let out = run_prints(
        r#"<?php
class NewsPublisher implements SplSubject {
    private SplObjectStorage $observers;
    public string $latestArticle = "";

    public function __construct() {
        $this->observers = new SplObjectStorage();
    }

    public function attach(SplObserver $observer): void {
        $this->observers->attach($observer);
    }

    public function detach(SplObserver $observer): void {
        $this->observers->detach($observer);
    }

    public function notify(): void {
        foreach ($this->observers as $observer) {
            $observer->update($this);
        }
    }

    public function publish(string $article): void {
        $this->latestArticle = $article;
        $this->notify();
    }
}

class Reader implements SplObserver {
    public string $received = "";
    public function update(SplSubject $subject): void {
        if ($subject instanceof NewsPublisher) {
            $this->received = $subject->latestArticle;
        }
    }
}

$publisher = new NewsPublisher();
$reader = new Reader();
$publisher->attach($reader);

$publisher->publish("PHP 8.4 Released!");
echo "Reader received: {$reader->received}";
"#,
    );
    assert_eq!(out, vec!["Reader received: PHP 8.4 Released!"]);
}

#[test]
fn test_php_spl_object_storage_contains_and_offset() {
    let out = run_prints(
        r#"<?php
$storage = new SplObjectStorage();
$o1 = new stdClass();
$o2 = new stdClass();

$storage->attach($o1, "data1");
$storage->attach($o2, "data2");

echo ($storage->contains($o1) ? "1" : "0") . " | data=" . $storage[$o1];
"#,
    );
    assert_eq!(out, vec!["1 | data=data1"]);
}

#[test]
fn test_php_spl_subject_detach_stops_notifications() {
    let out = run_prints(
        r#"<?php
class Emitter implements SplSubject {
    public SplObjectStorage $obs;
    public int $count = 0;
    public function __construct() { $this->obs = new SplObjectStorage(); }
    public function attach(SplObserver $o): void { $this->obs->attach($o); }
    public function detach(SplObserver $o): void { $this->obs->detach($o); }
    public function notify(): void { foreach ($this->obs as $o) { $o->update($this); } }
}

class Listener implements SplObserver {
    public int $events = 0;
    public function update(SplSubject $s): void { $this->events++; }
}

$emitter = new Emitter();
$l = new Listener();
$emitter->attach($l);
$emitter->notify();

$emitter->detach($l);
$emitter->notify();

echo "Received events: {$l->events}";
"#,
    );
    assert_eq!(out, vec!["Received events: 1"]);
}

#[test]
fn test_php_spl_object_storage_remove_all_and_add_all() {
    compile_ok(
        r#"<?php
$s1 = new SplObjectStorage();
$s2 = new SplObjectStorage();
$o1 = new stdClass(); $o2 = new stdClass();

$s1->attach($o1);
$s1->attach($o2);
$s2->addAll($s1);

echo count($s2) === 2 ? "ADD_ALL_OK" : "FAIL";
$s2->removeAll($s1);
echo count($s2) === 0 ? " REMOVE_ALL_OK" : " FAIL";
"#,
    );
}

#[test]
fn test_php_spl_object_storage_get_hash_custom_key() {
    compile_ok(
        r#"<?php
class CustomHashStorage extends SplObjectStorage {
    public function getHash(object $object): string {
        return spl_object_hash($object);
    }
}

$chs = new CustomHashStorage();
$o = new stdClass();
$chs->attach($o);
echo $chs->contains($o) ? "CUSTOM_HASH_OK" : "FAIL";
"#,
    );
}

#[test]
fn test_php_spl_object_storage_array_access_unsetting() {
    compile_ok(
        r#"<?php
$s = new SplObjectStorage();
$o = new stdClass();
$s[$o] = "value";
unset($s[$o]);
echo !isset($s[$o]) ? "UNSET_OK" : "FAIL";
"#,
    );
}

#[test]
fn test_php_spl_observer_multiple_listeners() {
    compile_ok(
        r#"<?php
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
"#,
    );
}

#[test]
fn test_php_spl_object_storage_serialize_unserialize() {
    compile_ok(
        r#"<?php
$s = new SplObjectStorage();
$o = new stdClass(); $o->name = "test";
$s->attach($o, "payload");

$serialized = serialize($s);
$restored = unserialize($serialized);
echo count($restored) === 1 ? "SERIALIZE_OK" : "FAIL";
"#,
    );
}

#[test]
fn test_php_spl_object_storage_iteration_rewind_valid() {
    compile_ok(
        r#"<?php
$s = new SplObjectStorage();
$s->attach(new stdClass(), 1);
$s->attach(new stdClass(), 2);

$s->rewind();
$count = 0;
while ($s->valid()) {
    $count++;
    $s->next();
}
echo "Iterated $count items";
"#,
    );
}

#[test]
fn test_php_spl_subject_weak_reference_storage() {
    compile_ok(
        r#"<?php
$storage = new SplObjectStorage();
$obj = new stdClass();
$storage->attach($obj);
echo $storage->contains($obj) ? "STORAGE_OK" : "FAIL";
"#,
    );
}
