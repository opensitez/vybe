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
}
