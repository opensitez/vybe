<?php
// vybe-test: php/weak_references/weak_map_metadata
// origin: languages/php/tests/php/test_weak_references.rs
// vybe-test-mode: compile

class Connection {
    public function __construct(public readonly string $dsn) {}
}
$map = new WeakMap();
$conn1 = new Connection('sqlite::memory:');
$conn2 = new Connection('mysql://localhost');
$map[$conn1] = ['created' => time(), 'queries' => 0];
$map[$conn2] = ['created' => time(), 'queries' => 0];
$map[$conn1]['queries']++;
echo $map[$conn1]['queries'];
echo count($map);
