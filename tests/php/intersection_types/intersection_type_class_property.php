<?php
// vybe-test: php/intersection_types/intersection_type_class_property
// origin: languages/php/tests/php/test_intersection_types.rs
// vybe-test-mode: compile

interface Closeable { public function close(): void; }
interface Flushable { public function flush(): void; }
class Buffer implements Closeable, Flushable {
    private string $data = '';
    public function write(string $s): void { $this->data .= $s; }
    public function flush(): void { echo $this->data; $this->data = ''; }
    public function close(): void { $this->flush(); }
}
class Writer {
    public Closeable&Flushable $buffer;
    public function __construct(Closeable&Flushable $buf) { $this->buffer = $buf; }
}
$w = new Writer(new Buffer());
