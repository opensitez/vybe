<?php
// vybe-test: php/generators_advanced/generator_return_type_in_interface
// origin: languages/php/tests/php/test_generators_advanced.rs
// vybe-test-mode: compile

interface Iterable2 {
    public function items(): Generator;
}
class NumberList implements Iterable2 {
    private array $data;
    public function __construct(array $data) { $this->data = $data; }
    public function items(): Generator {
        foreach ($this->data as $v) {
            yield $v;
        }
    }
}
$list = new NumberList([10, 20, 30]);
foreach ($list->items() as $v) {
    echo $v;
}
