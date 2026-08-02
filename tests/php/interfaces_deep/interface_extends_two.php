<?php
// vybe-test: php/interfaces_deep/interface_extends_two
// origin: languages/php/tests/php/test_interfaces_deep.rs
// vybe-test-mode: compile

interface Readable  { public function read(): string; }
interface Writable  { public function write(string $data): void; }
interface ReadWrite extends Readable, Writable {}
class File implements ReadWrite {
    private string $buffer = '';
    public function read(): string { return $this->buffer; }
    public function write(string $data): void { $this->buffer .= $data; }
}
$f = new File();
$f->write('hello');
$f->write(' world');
echo $f->read();
