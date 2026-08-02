<?php
// vybe-test: php/class_inspection/class_implements_interfaces_list
// origin: languages/php/tests/php/test_class_inspection.rs
// vybe-test-mode: compile

interface Serializable { public function serialize(): string; }
interface Loggable { public function log(): void; }
class Entity implements Serializable, Loggable {
    public function serialize(): string { return ''; }
    public function log(): void {}
}
$ifaces = class_implements('Entity');
echo isset($ifaces['Serializable']) ? 'yes' : 'no';
echo isset($ifaces['Loggable']) ? 'yes' : 'no';
