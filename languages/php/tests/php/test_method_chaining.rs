//! Fluent interfaces and `return $this` / `return static` method chains.

crate::php_cases! {
    chain_accumulates_string_parts => {
        r#"<?php
class Text {
    private string $buf = '';
    public function part(string $s): static {
        $this->buf .= $s;
        return $this;
    }
    public function done(): string { return $this->buf; }
}
echo (new Text())->part('a')->part('-')->part('b')->done();
"#,
        ["a-b"]
    };

    chain_static_factory_then_instance_methods => {
        r#"<?php
class Config {
    private array $data = [];
    public static function make(): static { return new static(); }
    public function set(string $k, mixed $v): static { $this->data[$k] = $v; return $this; }
    public function get(string $k): mixed { return $this->data[$k] ?? null; }
}
echo Config::make()->set('port', 8080)->get('port');
"#,
        ["8080"]
    };

    chain_subclass_returns_static_type => {
        r#"<?php
class Base {
    protected string $tag = '';
    public function mark(string $t): static { $this->tag = $t; return $this; }
    public function read(): string { return $this->tag; }
}
class Derived extends Base {}
echo (new Derived())->mark('d')->read();
"#,
        ["d"]
    };

    chain_through_magic_call_delegates => {
        r#"<?php
class Proxy {
    private int $n = 0;
    public function __call(string $name, array $args): static {
        if ($name === 'inc') { $this->n += (int)($args[0] ?? 1); }
        return $this;
    }
    public function val(): int { return $this->n; }
}
echo (new Proxy())->inc(2)->inc(3)->val();
"#,
        ["5"]
    };

    chain_pipe_applies_callables_in_order => {
        r#"<?php
class Pipe {
    public function __construct(private mixed $value) {}
    public function through(callable $fn): static {
        $this->value = $fn($this->value);
        return $this;
    }
    public function out(): mixed { return $this->value; }
}
echo (new Pipe(4))->through(fn($n) => $n * 2)->through(fn($n) => $n + 1)->out();
"#,
        ["9"]
    };

    chain_headers_then_body_on_response_builder => {
        r#"<?php
class Response {
    private array $headers = [];
    private string $body = '';
    public function header(string $k, string $v): static {
        $this->headers[] = "$k:$v";
        return $this;
    }
    public function body(string $b): static { $this->body = $b; return $this; }
    public function summary(): string {
        return count($this->headers) . '|' . $this->body;
    }
}
echo (new Response())->header('X', '1')->header('Y', '2')->body('ok')->summary();
"#,
        ["2|ok"]
    };

    chain_on_anonymous_class => {
        r#"<?php
$builder = new class {
    private array $items = [];
    public function push(string $v): static { $this->items[] = $v; return $this; }
    public function join(string $sep): string { return implode($sep, $this->items); }
};
echo $builder->push('x')->push('y')->join('-');
"#,
        ["x-y"]
    };

    chain_terminates_with_array_access_expression => {
        r#"<?php
class Rows {
    private array $rows = [];
    public function add(array $row): static { $this->rows[] = $row; return $this; }
    public function rows(): array { return $this->rows; }
}
echo (new Rows())->add(['id' => 1])->add(['id' => 2])->rows()[1]['id'];
"#,
        ["2"]
    };

    chain_interleaved_with_string_cast => {
        r#"<?php
class Label {
    private string $text = '';
    public function append(string $s): static { $this->text .= $s; return $this; }
    public function __toString(): string { return $this->text; }
}
echo (string)(new Label())->append('hi')->append('!');
"#,
        ["hi!"]
    };

    chain_parent_reference_after_child_chain => {
        r#"<?php
class Node {
    public function __construct(public string $name) {}
    public function rename(string $n): static { $this->name = $n; return $this; }
}
$n = (new Node('a'))->rename('b');
echo $n->name;
"#,
        ["b"]
    };

    chain_with_interface_return_type => {
        r#"<?php
interface Step { public function bump(): self; public function value(): int; }
class Counter implements Step {
    private int $n = 0;
    public function bump(): self { $this->n++; return $this; }
    public function value(): int { return $this->n; }
}
/** @var Step $c */
$c = new Counter();
echo $c->bump()->bump()->value();
"#,
        ["2"]
    };

    chain_nested_calls_share_same_instance_state => {
        r#"<?php
class Stack {
    private array $s = [];
    public function push(int $v): static { $this->s[] = $v; return $this; }
    public function top(): int { return $this->s[count($this->s) - 1]; }
}
$st = new Stack();
$st->push(1)->push(2);
echo $st->top();
"#,
        ["2"]
    };

    chain_bool_flags_then_render => {
        r#"<?php
class Flags {
    private int $mask = 0;
    public function on(int $bit): static { $this->mask |= $bit; return $this; }
    public function has(int $bit): bool { return ($this->mask & $bit) !== 0; }
}
$f = (new Flags())->on(1)->on(4);
echo ($f->has(1) ? 'a' : '') . ($f->has(4) ? 'b' : '');
"#,
        ["ab"]
    };

    chain_with_null_coalescing_fallback_after_terminal => {
        r#"<?php
class Maybe {
    private ?string $v = null;
    public function set(string $s): static { $this->v = $s; return $this; }
    public function get(): ?string { return $this->v; }
}
echo (new Maybe())->set('z')->get() ?? 'none';
"#,
        ["z"]
    };

    chain_deep_five_levels_same_line => {
        r#"<?php
class N {
    private int $v = 0;
    public function add(int $x): static { $this->v += $x; return $this; }
    public function val(): int { return $this->v; }
}
echo (new N())->add(1)->add(2)->add(3)->add(4)->add(5)->val();
"#,
        ["15"]
    };

    chain_through_array_map_closure => {
        r#"<?php
class ListBuilder {
    private array $items = [];
    public function add(int $v): static { $this->items[] = $v; return $this; }
    public function map(callable $fn): static { $this->items = array_map($fn, $this->items); return $this; }
    public function first(): int { return $this->items[0]; }
}
echo (new ListBuilder())->add(1)->add(2)->map(fn($n) => $n * 10)->first();
"#,
        ["10"]
    };

    chain_with_private_method_via_public_proxy => {
        r#"<?php
class Builder {
    private int $n = 0;
    private function bump(): void { $this->n++; }
    public function step(): static { $this->bump(); return $this; }
    public function value(): int { return $this->n; }
}
echo (new Builder())->step()->step()->step()->value();
"#,
        ["3"]
    };

    chain_with_fluent_trait_methods => {
        r#"<?php
trait Fluent {
    public function setA(int $n): static { $this->a = $n; return $this; }
}
class Holder {
    use Fluent;
    public int $a = 0;
    public function setB(int $n): static { $this->a += $n; return $this; }
    public function value(): int { return $this->a; }
}
echo (new Holder())->setA(2)->setB(3)->value();
"#,
        ["5"]
    };

    chain_with_callables_and_array_access => {
        r#"<?php
class Bag {
    public array $items = [];
    public function push(string $v): static { $this->items[] = $v; return $this; }
}
$bag = (new Bag())->push('a')->push('b');
$bag->items[] = 'c';
echo implode('-', $bag->items);
"#,
        ["a-b-c"]
    };

    chain_with_conditional_method_skip => {
        r#"<?php
class Condition {
    private int $n = 0;
    public function add(int $x): static { $this->n += $x; return $this; }
    public function maybe(bool $ok): ?Condition {
        return $ok ? $this : null;
    }
    public function get(): int { return $this->n; }
}
echo (new Condition())->add(1)->maybe(true)->add(2)->get();
"#,
        ["3"]
    };

    chain_on_clone_does_not_mutate_original => {
        r#"<?php
class Cursor {
    public function __construct(private int $v) {}
    public function inc(int $d): static { $this->v += $d; return $this; }
    public function value(): int { return $this->v; }
}
$a = new Cursor(1);
$b = (clone $a)->inc(4);
echo $a->value() . '|' . $b->value();
"#,
        ["1|5"]
    };

    chain_with_nullsafe_operator_and_fallback => {
        r#"<?php
class MaybeChain {
    public function step(): static {
        return $this;
    }
    public function value(): int {
        return 42;
    }
}

$obj = new MaybeChain();
$present = $obj?->step()?->value();
$absent = null?->step()?->value();
echo ($present === 42 ? 'yes' : 'no') . '|' . ($absent === null ? 'null' : 'val');
"#,
        ["yes|null"]
    };

    chain_when_step_returns_new_instance => {
        r#"<?php
class Numberer {
    private int $v;
    public function __construct(int $v = 0) { $this->v = $v; }
    public function withOffset(int $d): Numberer { return new Numberer($this->v + $d); }
    public function value(): int { return $this->v; }
}
$n = (new Numberer())->withOffset(4)->withOffset(5)->value();
echo $n;
"#,
        ["9"]
    };

    chain_with_void_step_then_tail_method => {
        r#"<?php
class Pipeline {
    private string $log = '';
    public function touch(string $chunk): void {
        $this->log .= $chunk;
    }
    public function chain(string $chunk): static {
        $this->touch($chunk);
        return $this;
    }
    public function snapshot(): string { return $this->log; }
}
echo (new Pipeline())->chain('a')->chain('b')->snapshot();
"#,
        ["ab"]
    };

    chain_with_intermixed_clone_points => {
        r#"<?php
class Counter {
    public function __construct(private int $v) {}
    public function with(int $d): Counter { return new Counter($this->v + $d); }
    public function plus(int $d): self { $this->v += $d; return $this; }
    public function value(): int { return $this->v; }
}
$base = new Counter(1);
$copy = $base->with(4)->plus(3);
echo $base->value() . '|' . $copy->value();
"#,
        ["1|8"]
    };

    chain_with_guarded_call_then_unwrap => {
        r#"<?php
class Guard {
    private int $n = 0;
    public function add(int $v): static { $this->n += $v; return $this; }
    public function maybe(bool $ok): static {
        return $ok ? $this : $this;
    }
    public function total(): int { return $this->n; }
}
echo (new Guard())->add(1)->maybe(true)->add(2)->total();
"#,
        ["3"]
    };

    chain_nullsafe_followed_by_terminal => {
        r#"<?php
class Item {
    public function step(): self { return $this; }
    public function id(): string { return 'id'; }
}
$item = new Item();
$chain = $item?->step()?->id();
echo $chain;
"#,
        ["id"]
    };

    chain_with_nested_anonymous_class => {
        r#"<?php
$builder = new class {
    private array $parts = [];
    public function add(string $s): static { $this->parts[] = $s; return $this; }
    public function chain(int $i): self {
        $this->parts[] = (string)$i;
        return $this;
    }
    public function join(): string { return implode('|', $this->parts); }
};
echo $builder->add('x')->chain(9)->join();
"#,
        ["x|9"]
    };

    chain_with_static_factory_and_fluent_access => {
        r#"<?php
class Service {
    private array $events = [];
    public static function from(string $first): static {
        return new static($first);
    }
    private function __construct(private string $seed) {}
    public function append(string $part): static {
        $this->events[] = $part;
        return $this;
    }
    public function summary(): string {
        return $this->seed . ':' . implode('+', $this->events);
    }
}
echo Service::from('root')->append('a')->append('b')->summary();
"#,
        ["root:a+b"]
    };

    chain_with_trait_fluent_methods => {
        r#"<?php
trait TChain {
    public function add(int $v): static {
        if (!isset($this->items)) {
            $this->items = [];
        }
        $this->items[] = $v;
        return $this;
    }
}
class Basket {
    use TChain;
    public array $items = [];
    public function total(): int { return array_sum($this->items); }
}
echo (new Basket())->add(2)->add(3)->total();
"#,
        ["5"]
    };

    chain_with_conditional_ternary_and_method_chain => {
        r#"<?php
class Flagger {
    private int $n = 0;
    public function add(int $x): static { $this->n += $x; return $this; }
    public function value(): int { return $this->n; }
}
$enabled = true;
$result = $enabled ? (new Flagger())->add(1)->add(2) : (new Flagger());
echo $result->value();
"#,
        ["3"]
    };

    chain_with_intermediate_object_reassignment => {
        r#"<?php
class Stage {
    private int $v = 0;
    public function add(int $x): static { $this->v += $x; return $this; }
    public function fork(): static {
        $next = new Stage();
        $next->add($this->v);
        return $next;
    }
    public function value(): int { return $this->v; }
}
$start = new Stage();
$final = $start->add(3)->fork()->add(4);
echo $start->value() . '|' . $final->value();
"#,
        ["0|7"]
    };

    chain_with_final_method_on_previous_link => {
        r#"<?php
class Link {
    public function __construct(private int $n) {}
    public function inc(int $x): static { $this->n += $x; return $this; }
    public function next(): static { return $this; }
    public function value(): int { return $this->n; }
}
echo (new Link(1))->inc(2)->next()->inc(3)->value();
"#,
        ["6"]
    };

    chain_with_trait_and_static_return => {
        r#"<?php
trait FluentMath {
    public function addOne(): static {
        $this->value += 1;
        return $this;
    }
}
class Counter {
    use FluentMath;
    public int $value = 0;
    public function add(int $v): static {
        $this->value += $v;
        return $this;
    }
}
echo (new Counter())->add(2)->addOne()->add(3)->addOne()->value;
"#,
        ["7"]
    };

    chain_with_callable_stepper => {
        r#"<?php
class Stepper {
    private int $x = 0;
    public function apply(callable $fn): static {
        $this->x = $fn($this->x);
        return $this;
    }
    public function value(): int { return $this->x; }
}
echo (new Stepper())->apply(fn($v) => $v + 1)->apply(fn($v) => $v * 4)->value();
"#,
        ["4"]
    };

    chain_with_method_and_array_key_access => {
        r#"<?php
class Record {
    private array $items = [];
    public function set(string $k, int $v): static { $this->items[$k] = $v; return $this; }
    public function getAll(): array { return $this->items; }
}
$record = (new Record())->set('first', 1)->set('second', 2);
echo $record->getAll()['second'];
"#,
        ["2"]
    };

    chain_with_magic_call_proxy => {
        r#"<?php
class Proxy {
    private int $value = 0;
    public function __call(string $name, array $args): static {
        if ($name === 'add') { $this->value += (int)($args[0] ?? 0); }
        return $this;
    }
    public function value(): int { return $this->value; }
}
echo (new Proxy())->add(2)->add(3)->value();
"#,
        ["5"]
    };

    chain_with_dynamic_method_name => {
        r#"<?php
class Writer {
    private string $s = '';
    public function write(string $chunk): static { $this->s .= $chunk; return $this; }
    public function output(): string { return $this->s; }
}
$writer = new Writer();
$method = 'write';
echo $writer->{$method}('a')->{$method}('b')->output();
"#,
        ["ab"]
    };

    chain_with_fluent_clone_copy => {
        r#"<?php
class Token {
    public function __construct(private int $n = 0) {}
    public function add(int $d): static { $this->n += $d; return $this; }
    public function value(): int { return $this->n; }
}
$base = new Token(1);
$next = clone $base;
echo $base->add(2)->value() . '|' . $next->add(3)->value();
"#,
        ["3|4"]
    };

    chain_with_intermediate_snapshot => {
        r#"<?php
class State {
    private int $n = 0;
    public function inc(int $v): static { $this->n += $v; return $this; }
    public function snapshot(): int { return $this->n; }
    public function value(): int { return $this->n; }
}
$s = new State();
$before = $s->inc(1)->inc(1)->snapshot();
$after = $s->inc(3)->value();
echo $before . '|' . $after;
"#,
        ["2|5"]
    };

    chain_with_callback_chain => {
        r#"<?php
class Pipeline {
    private int $n = 0;
    public function then(callable $fn): static { $this->n = $fn($this->n); return $this; }
    public function value(): int { return $this->n; }
}
echo (new Pipeline())->then(fn(int $n) => $n + 2)->then(fn(int $n) => $n * 3)->then(fn(int $n) => $n - 1)->value();
"#,
        ["5"]
    };

    chain_with_nested_chain_return_self => {
        r#"<?php
class Logger {
    private array $events = [];
    public function push(string $e): static { $this->events[] = $e; return $this; }
    public function child(): static { return $this; }
    public function count(): int { return count($this->events); }
}
$log = new Logger();
echo $log->push('a')->push('b')->child()->push('c')->count();
"#,
        ["3"]
    };

    chain_with_return_self_variants => {
        r#"<?php
class Builder {
    private array $steps = [];
    public function set(string $name, int $value): static { $this->steps[$name] = $value; return $this; }
    public function merge(array $data): static { $this->steps = array_merge($this->steps, $data); return $this; }
    public function size(): int { return count($this->steps); }
}
echo (new Builder())
    ->set('a', 1)
    ->merge(['b' => 2])
    ->set('c', 3)
    ->size();
"#,
        ["3"]
    };

chain_with_trait_return_static => {
        r#"<?php
trait Fluent {
    public array $tags = [];
    public function add(string $name): static { $this->tags[] = $name; return $this; }
}
class Tagger {
    use Fluent;
    public function count(): int { return count($this->tags); }
}
    echo (new Tagger())->add('x')->add('y')->count();
"#,
        ["2"]
    };
    chain_dynamic_method_name_with_call_chain_and_arguments => {
        r#"<?php
class Writer {
    private string $s = '';
    public function append(string $chunk): static { $this->s .= $chunk; return $this; }
    public function output(): string { return $this->s; }
}
$writer = new Writer();
$method = 'append';
echo $writer->{$method}('a')->{$method}('b')->output();
"#,
        ["ab"]
    };

    chain_array_access_after_dynamic_property_chain => {
        r#"<?php
class Store {
    private array $buckets = [];
    public function bucket(int $idx): static {
        if (!isset($this->buckets[$idx])) { $this->buckets[$idx] = []; }
        return $this;
    }
    public function put(int $idx, string $value): static {
        $this->buckets[$idx][] = $value;
        return $this;
    }
    public function buckets(): array { return $this->buckets; }
}
echo (new Store())->bucket(1)->put(1, 'x')->buckets()[1][0];
"#,
        ["x"]
    };

    chain_operator_precedence_with_parentheses => {
        r#"<?php
class Counter {
    private int $n = 0;
    public function inc(int $d): static { $this->n += $d; return $this; }
    public function value(): int { return $this->n; }
}
echo ((new Counter())->inc(1)->value() + 1) . '|';
echo ((new Counter())->inc(2)->value()) + (1 + 1);
"#
        ,
        ["2|4"]
    };

    chain_with_conditional_method_chain => {
        r#"<?php
class Guarded {
    private int $n = 0;
    public function add(int $x): static { $this->n += $x; return $this; }
    public function maybe(bool $ok): ?self { return $ok ? $this : null; }
    public function value(): int { return $this->n; }
}
echo ($ok = (new Guarded())->add(1)->maybe(true)->add(2)->value());
echo '|' . ((new Guarded())->maybe(false)?->add(3)?->value() ?? 99);
"#,
        ["3|99"]
    };

    chain_with_clone_mid_chain => {
        r#"<?php
class Draft {
    private int $n = 0;
    public function inc(int $x): static { $this->n += $x; return $this; }
    public function value(): int { return $this->n; }
}
$base = new Draft();
$clone = clone $base;
$base->inc(4);
$clone->inc(1)->inc(2);
echo $base->value() . '|' . $clone->value();
"#,
        ["4|3"]
    };

    chain_with_property_accessor_chain => {
        r#"<?php
class Payload {
    public function __construct(private array $state = []) {}
    public function put(string $k, string $v): static {
        $this->state[$k] = $v;
        return $this;
    }
    public function state(): array { return $this->state; }
}
echo (new Payload())->put('a', 'x')->put('b', 'y')->state()['b'];
"#,
        ["y"]
    };

    chain_with_method_returning_array => {
        r#"<?php
class Collector {
    private array $items = [];
    public function add(string $v): static { $this->items[] = $v; return $this; }
    public function all(): array { return $this->items; }
}
echo implode('-', (new Collector())->add('p')->add('q')->all());
"#,
        ["p-q"]
    };
}
