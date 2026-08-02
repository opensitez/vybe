<?php
// vybe-test: php/oop_patterns/multiple_traits_methods_from_each
// origin: languages/php/tests/php/test_oop_patterns.rs
// vybe-test-mode: compile

trait Serializable {
    public function serialize(): string { return json_encode($this->toArray()); }
}
trait Loggable {
    public function log(): string { return get_class($this) . ':logged'; }
}
class Order {
    use Serializable, Loggable;
    public function __construct(private int $id, private float $total) {}
    public function toArray(): array { return ['id' => $this->id, 'total' => $this->total]; }
}
$o = new Order(42, 99.99);
echo $o->log();
echo is_string($o->serialize()) ? 'serialized' : 'fail';
