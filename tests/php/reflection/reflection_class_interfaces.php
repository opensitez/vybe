<?php
// vybe-test: php/reflection/reflection_class_interfaces
// origin: languages/php/tests/php/test_reflection.rs
// vybe-test-mode: compile

interface Serializable2 { public function serialize2(): string; }
interface Loggable { public function log(): void; }
class Service implements Serializable2, Loggable {
    public function serialize2(): string { return ''; }
    public function log(): void {}
}
$rc = new ReflectionClass(Service::class);
$ifaces = $rc->getInterfaceNames();
sort($ifaces);
echo implode(',', $ifaces);
