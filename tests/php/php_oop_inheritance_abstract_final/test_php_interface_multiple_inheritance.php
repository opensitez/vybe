<?php
// vybe-test: php/php_oop_inheritance_abstract_final/test_php_interface_multiple_inheritance
// origin: languages/php/tests/php/test_php_oop_inheritance_abstract_final.rs
// vybe-test-mode: compile

interface Readable { public function read(): string; }
interface Writable { public function write(string $data): void; }
interface ReadWriteable extends Readable, Writable {}

class Buffer implements ReadWriteable {
    private string $content = "";
    public function read(): string { return $this->content; }
    public function write(string $data): void { $this->content .= $data; }
}

    $b = new Buffer();
$b->write("hello");
echo $b->read();
