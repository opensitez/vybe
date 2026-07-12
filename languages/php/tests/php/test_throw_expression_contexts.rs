//! PHP 8 throw expressions in distinct syntactic positions (not catch-handler padding).

crate::php_cases! {
    throw_in_null_coalesce_left_missing => {
        r#"<?php
$v = null;
try { echo $v ?? throw new RuntimeException('missing'); }
catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["missing"]
    };

    throw_in_null_coalesce_chain_middle => {
        r#"<?php
$a = null;
$b = null;
try { echo $a ?? $b ?? throw new LogicException('chain'); }
catch (LogicException $e) { echo $e->getMessage(); }
"#,
        ["chain"]
    };

    throw_in_elvis_operator_falsy => {
        r#"<?php
$flag = 0;
try { echo $flag ?: throw new InvalidArgumentException('falsy'); }
catch (InvalidArgumentException $e) { echo $e->getMessage(); }
"#,
        ["falsy"]
    };

    throw_in_match_default_arm => {
        r#"<?php
try {
    echo match (3) { 1 => 'one', 2 => 'two', default => throw new DomainException('nomatch') };
} catch (DomainException $e) { echo $e->getMessage(); }
"#,
        ["nomatch"]
    };

    throw_in_match_conditional_arm => {
        r#"<?php
try {
    echo match (true) {
        true => throw new OverflowException('hot'),
        false => 'cold',
    };
} catch (OverflowException $e) { echo $e->getMessage(); }
"#,
        ["hot"]
    };

    throw_in_short_arrow_fn_body => {
        r#"<?php
$need = fn(?string $s) => $s ?? throw new ValueError('empty');
try { $need(null); } catch (ValueError $e) { echo $e->getMessage(); }
"#,
        ["empty"]
    };

    throw_in_long_arrow_fn_expression => {
        r#"<?php
$pick = fn(array $xs) => count($xs) > 0 ? $xs[0] : throw new UnderflowException('no head');
try { $pick([]); } catch (UnderflowException $e) { echo $e->getMessage(); }
"#,
        ["no head"]
    };

    throw_as_function_argument => {
        r#"<?php
function wrap(string $msg): string { return "[$msg]"; }
try { echo wrap(throw new Exception('arg')); }
catch (Exception $e) { echo $e->getMessage(); }
"#,
        ["arg"]
    };

    throw_in_array_literal_value => {
        r#"<?php
try {
    $pair = ['ok' => 1, 'bad' => throw new RuntimeException('elem')];
} catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["elem"]
    };

    throw_in_return_statement => {
        r#"<?php
function boom(): never {
    throw new Exception('ret');
}
try { boom(); } catch (Exception $e) { echo $e->getMessage(); }
"#,
        ["ret"]
    };

    throw_in_property_assignment_via_coalesce => {
        r#"<?php
class Holder { public ?string $slot = null; }
$h = new Holder();
try {
    $h->slot = $h->slot ?? throw new LogicException('unset');
} catch (LogicException $e) { echo $e->getMessage(); }
"#,
        ["unset"]
    };

    throw_in_echo_statement_expression => {
        r#"<?php
try { echo throw new RuntimeException('echo'); }
catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["echo"]
    };

    throw_in_nested_match_inner_arm => {
        r#"<?php
try {
    $outer = match ('go') {
        'go' => match (0) { 1 => 'hit', default => throw new RuntimeException('inner') },
        default => 'skip',
    };
    echo $outer;
} catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["inner"]
    };

    throw_in_list_destructure_default => {
        r#"<?php
try {
    [$a, $b] = [1, throw new ValueError('rhs')];
} catch (ValueError $e) { echo $e->getMessage(); }
"#,
        ["rhs"]
    };

    throw_in_foreach_value_position => {
        r#"<?php
$log = [];
foreach ([1, 0, 2] as $n) {
    try {
        $log[] = $n ?: throw new RuntimeException("zero:$n");
    } catch (RuntimeException $e) {
        $log[] = $e->getMessage();
    }
}
echo implode('|', $log);
"#,
        ["1|zero:0|2"]
    };

    throw_in_concatenation_rhs => {
        r#"<?php
$prefix = 'err:';
try { echo $prefix . (throw new Exception('tail')); }
catch (Exception $e) { echo $e->getMessage(); }
"#,
        ["tail"]
    };

    throw_in_bitwise_or_fallback => {
        r#"<?php
$mask = 0;
try { echo $mask | throw new LogicException('or'); }
catch (LogicException $e) { echo $e->getMessage(); }
"#,
        ["or"]
    };

    throw_in_spaceship_compare_operand => {
        r#"<?php
try {
    $a = 1;
    $b = throw new RuntimeException('cmp');
} catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["cmp"]
    };

    throw_in_instanceof_guard => {
        r#"<?php
$obj = new stdClass();
try {
    echo ($obj instanceof Stringable) ? 'yes' : throw new TypeError('not stringable');
} catch (TypeError $e) { echo $e->getMessage(); }
"#,
        ["not stringable"]
    };

    throw_in_nullsafe_chain_fallback => {
        r#"<?php
class Node { public ?Node $next = null; }
$root = new Node();
try {
    $v = $root?->next?->next ?? throw new RuntimeException('deep');
} catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["deep"]
    };

    throw_in_yield_from_expression => {
        r#"<?php
function bad(): Generator {
    yield from (throw new Exception('yieldfrom'));
}
try { iterator_to_array(bad()); } catch (Exception $e) { echo $e->getMessage(); }
"#,
        ["yieldfrom"]
    };

    throw_in_generator_send_result => {
        r#"<?php
function gen(): Generator {
  $x = yield 'step';
  echo $x ?? throw new RuntimeException('no send');
}
$g = gen();
echo $g->current();
$g->send(null);
try { $g->send('ok'); } catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["step", "no send"]
    };

    throw_in_switch_default_via_match => {
        r#"<?php
$code = 9;
try {
    echo match ($code) { 1 => 'a', 2 => 'b', default => throw new UnexpectedValueException('code') };
} catch (UnexpectedValueException $e) { echo $e->getMessage(); }
"#,
        ["code"]
    };

    throw_in_class_constant_via_method => {
        r#"<?php
class Config {
    public static function port(): int {
        return $_ENV['PORT'] ?? throw new RuntimeException('no port');
    }
}
try { echo Config::port(); } catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["no port"]
    };

    throw_in_interface_default_method => {
        r#"<?php
interface Parser {
    public function parse(string $raw): array;
    public function requireNonEmpty(?string $raw): string {
        return $raw ?? throw new InvalidArgumentException('blank');
    }
}
class JsonParser implements Parser {
    public function parse(string $raw): array { return []; }
}
$p = new JsonParser();
try { $p->requireNonEmpty(null); } catch (InvalidArgumentException $e) { echo $e->getMessage(); }
"#,
        ["blank"]
    };

    throw_in_trait_method_body => {
        r#"<?php
trait Guard {
    public function requirePositive(int $n): int {
        return $n > 0 ? $n : throw new ValueError('nonpositive');
    }
}
class Counter { use Guard; }
$c = new Counter();
try { $c->requirePositive(0); } catch (ValueError $e) { echo $e->getMessage(); }
"#,
        ["nonpositive"]
    };

    throw_in_enum_method => {
        r#"<?php
enum Status: string {
    case On = 'on';
    case Off = 'off';
    public function label(): string {
        return match ($this) {
            self::On => 'enabled',
            self::Off => throw new LogicException('hidden'),
        };
    }
}
try { echo Status::Off->label(); } catch (LogicException $e) { echo $e->getMessage(); }
"#,
        ["hidden"]
    };

    throw_in_readonly_constructor_promotion_check => {
        r#"<?php
readonly class Token {
    public function __construct(public string $value) {
        if ($value === '') { throw new ValueError('empty token'); }
    }
}
try { new Token(''); } catch (ValueError $e) { echo $e->getMessage(); }
"#,
        ["empty token"]
    };

    throw_in_attribute_constructor => {
        r#"<?php
#[\Attribute]
class Route {
    public function __construct(public string $path) {
        if ($path[0] !== '/') { throw new InvalidArgumentException('path'); }
    }
}
try { new Route('bad'); } catch (InvalidArgumentException $e) { echo $e->getMessage(); }
"#,
        ["path"]
    };

    throw_in_splat_argument_position => {
        r#"<?php
function sum(int ...$nums): int { return array_sum($nums); }
$args = [1, throw new RuntimeException('splat')];
try { sum(...$args); } catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["splat"]
    };

    throw_in_new_expression_argument => {
        r#"<?php
class Box { public function __construct(public int $size) {} }
try { new Box(throw new ValueError('size')); } catch (ValueError $e) { echo $e->getMessage(); }
"#,
        ["size"]
    };

    throw_in_catch_rethrow_wrapper => {
        r#"<?php
try {
    try { throw new RuntimeException('inner'); }
    catch (RuntimeException $e) { throw $e ?? throw new LogicException('impossible'); }
} catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["inner"]
    };

    throw_in_finally_suppressed_by_catch_return => {
        r#"<?php
function f(): string {
    try { throw new Exception('x'); }
    catch (Exception $e) { return 'caught'; }
    finally { throw new RuntimeException('finally'); }
}
try { echo f(); } catch (RuntimeException $e) { echo 'late'; }
"#,
        ["late"]
    };

    throw_in_do_while_condition_via_assignment => {
        r#"<?php
$i = 0;
try {
    do {
        $i++;
        if ($i === 2) { throw new RuntimeException('loop'); }
    } while ($i < 1);
} catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["loop"]
    };

    throw_in_while_assignment_expression => {
        r#"<?php
$queue = [0];
try {
    while (($n = array_shift($queue)) !== null) {
        throw new RuntimeException("left:$n");
    }
} catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["left:0"]
    };

    throw_in_for_init_clause => {
        r#"<?php
try {
    for ($i = throw new ValueError('init'); $i < 1; $i++) { echo 'nope'; }
} catch (ValueError $e) { echo $e->getMessage(); }
"#,
        ["init"]
    };

    throw_in_for_condition_expression => {
        r#"<?php
try {
    for ($i = 0; $i < (throw new LogicException('cond')); $i++) { echo 'nope'; }
} catch (LogicException $e) { echo $e->getMessage(); }
"#,
        ["cond"]
    };

    throw_in_for_update_expression => {
        r#"<?php
$log = [];
try {
    for ($i = 0; $i < 2; $i = $i + (throw new RuntimeException('upd'))) {}
} catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["upd"]
    };

    throw_in_if_short_circuit_and => {
        r#"<?php
$ok = false;
try {
    echo ($ok && throw new RuntimeException('and')) ? 'yes' : 'no';
} catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["no"]
    };

    throw_in_if_short_circuit_or => {
        r#"<?php
$ok = true;
try {
    echo ($ok || throw new RuntimeException('or')) ? 'yes' : 'no';
} catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["yes"]
    };

    throw_in_ternary_nested_condition => {
        r#"<?php
$mode = 'strict';
try {
    echo $mode === 'strict' ? 'ok' : ($mode ?? throw new RuntimeException('mode'));
} catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["ok"]
    };

    throw_in_string_interpolation_expression => {
        r#"<?php
$name = null;
try { echo "hello {$name ?? throw new RuntimeException('name')}"; }
catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["name"]
    };

    throw_in_heredoc_interpolation => {
        r#"<?php
$id = null;
try {
    echo <<<TXT
id={$id ?? throw new RuntimeException('id')}
TXT;
} catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["id"]
    };

    throw_in_array_unpack_position => {
        r#"<?php
$head = [1];
try { $all = [...$head, ...(throw new RuntimeException('spread'))]; }
catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["spread"]
    };

    throw_in_named_argument_value => {
        r#"<?php
function pair(int $a, int $b): int { return $a + $b; }
try { echo pair(a: 1, b: throw new ValueError('b')); }
catch (ValueError $e) { echo $e->getMessage(); }
"#,
        ["b"]
    };

    throw_in_first_class_callable_invoke => {
        r#"<?php
class Math { public function inc(int $n): int { return $n + 1; } }
$fn = (new Math())->inc(...);
try { $fn(throw new ValueError('n')); } catch (ValueError $e) { echo $e->getMessage(); }
"#,
        ["n"]
    };

    throw_in_closure_use_by_reference_guard => {
        r#"<?php
$flag = false;
$run = function() use (&$flag): string {
    return $flag ? 'on' : throw new RuntimeException('off');
};
try { echo $run(); } catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["off"]
    };

    throw_in_static_variable_guard => {
        r#"<?php
function once(): string {
    static $done = false;
    return $done ? 'again' : ($done = true) || throw new RuntimeException('first');
}
try { echo once(); } catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["1"]
    };

    throw_in_match_with_multiple_conditions => {
        r#"<?php
$n = 15;
try {
    echo match (true) {
        $n < 10 => 'low',
        $n < 20 => throw new RangeException('mid'),
        default => 'high',
    };
} catch (RangeException $e) { echo $e->getMessage(); }
"#,
        ["mid"]
    };

    throw_in_union_type_coalesce_fallback => {
        r#"<?php
function asInt(string|int|null $v): int {
    return is_int($v) ? $v : (is_string($v) ? (int)$v : throw new TypeError('not scalar'));
}
try { echo asInt(null); } catch (TypeError $e) { echo get_class($e); }
"#,
        ["TypeError"]
    };

    throw_in_nullsafe_method_call_chain => {
        r#"<?php
class Api { public function token(): ?string { return null; } }
$api = new Api();
try {
    $t = $api?->token() ?? throw new RuntimeException('token');
} catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["token"]
    };

    throw_in_array_reduce_initial => {
        r#"<?php
$nums = [1, 2, 3];
try {
    array_reduce($nums, fn($c, $n) => $c + $n, throw new ValueError('seed'));
} catch (ValueError $e) { echo $e->getMessage(); }
"#,
        ["seed"]
    };

    throw_in_json_encode_option_guard => {
        r#"<?php
$data = ["\xB1\x31"];
$flags = JSON_THROW_ON_ERROR;
try {
    json_encode($data, $flags | (false ? 0 : throw new ValueError('flag')));
} catch (ValueError $e) { echo $e->getMessage(); }
"#,
        ["flag"]
    };

    throw_in_preg_replace_callback_return => {
        r#"<?php
try {
    preg_replace_callback('/\d+/', fn($m) => throw new RuntimeException('digit'), 'a1b');
} catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["digit"]
    };

    throw_in_array_map_callback => {
        r#"<?php
try {
    array_map(fn($v) => $v === 0 ? throw new ValueError('zero') : $v * 2, [1, 0]);
} catch (ValueError $e) { echo $e->getMessage(); }
"#,
        ["zero"]
    };

    throw_in_usort_comparator => {
        r#"<?php
$xs = [3, 1, 2];
try {
    usort($xs, fn($a, $b) => $a === $b ? 0 : throw new LogicException('cmp'));
} catch (LogicException $e) { echo $e->getMessage(); }
"#,
        ["cmp"]
    };

    throw_in_array_filter_callback => {
        r#"<?php
try {
    array_filter([1, 2], fn($v) => $v > 5 ? true : throw new RuntimeException('small'));
} catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["small"]
    };

    throw_in_serialize_allowed_classes_guard => {
        r#"<?php
class Secret {}
$blob = serialize(new Secret());
try {
    unserialize($blob, ['allowed_classes' => false ? [Secret::class] : throw new RuntimeException('deny')]);
} catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["deny"]
    };
}
