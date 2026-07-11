//! Generator `send`, `throw`, nested `yield from` — beyond `test_generators.rs`.

crate::php_cases! {
    generator_send_injects_value => {
        r#"<?php
function acc(): Generator {
    $total = 0;
    while (true) {
        $total += yield $total;
    }
}
$g = acc();
$g->current();
echo $g->send(3);
echo $g->send(4);
"#,
        ["37"]
    };

    generator_yield_from_generator => {
        r#"<?php
function inner(): Generator { yield 'a'; yield 'b'; }
function outer(): Generator { yield from inner(); yield 'c'; }
echo implode('', iterator_to_array(outer()));
"#,
        ["cb"]
    };

    generator_yield_from_array => {
        r#"<?php
function g(): Generator { yield from [1, 2]; yield 3; }
echo implode(',', iterator_to_array(g()));
"#,
        ["3,2"]
    };

    generator_keys_reset_after_yield_from => {
        r#"<?php
function g(): Generator { yield from ['x' => 1, 'y' => 2]; }
echo json_encode(iterator_to_array(g()));
"#,
        ["{\"x\":1,\"y\":2}"]
    };

    generator_return_value_after_yield_from => {
        r#"<?php
function inner(): Generator { yield 1; return 'done'; }
function outer(): Generator { $r = yield from inner(); return $r; }
$g = outer();
iterator_to_array($g);
echo $g->getReturn();
"#,
        ["done"]
    };

    generator_foreach_by_reference_key => {
        r#"<?php
function g(): Generator { yield 'k' => 'v'; }
foreach (g() as $k => $v) { echo $k . $v; }
"#,
        ["kv"]
    };

    generator_to_array_preserves_keys => {
        r#"<?php
function g(): Generator { yield 'a' => 1; yield 'b' => 2; }
echo count(g());
"#,
        ["2"]
    };

    generator_nested_yield_single => {
        r#"<?php
function g(): Generator { yield 1; yield 2; }
function wrap(): Generator { yield from g(); }
echo implode('', iterator_to_array(wrap()));
"#,
        ["12"]
    };

    generator_valid_true_before_first_next => {
        r#"<?php
function g(): Generator { yield 1; }
$g = g();
echo $g->valid() ? 'yes' : 'no';
"#,
        ["yes"]
    };

    generator_rewind_not_supported_still_valid => {
        r#"<?php
function g(): Generator { yield 1; }
$g = g();
echo $g->valid() ? '1' : '0';
"#,
        ["1"]
    };

    generator_yield_expression_in_expression => {
        r#"<?php
function g(): Generator { $x = yield 1; yield $x; }
$gen = g();
$gen->next();
echo $gen->send(9);
"#,
        ["9"]
    };

    generator_delegation_empty_inner => {
        r#"<?php
function empty_gen(): Generator { if (false) yield; return; }
function outer(): Generator { yield from empty_gen(); yield 'z'; }
echo implode('', iterator_to_array(outer()));
"#,
        ["z"]
    };

    generator_multiple_yield_from_chain => {
        r#"<?php
function a(): Generator { yield 1; }
function b(): Generator { yield from a(); yield 2; }
function c(): Generator { yield from b(); yield 3; }
echo implode('', iterator_to_array(c()));
"#,
        ["3"]
    };

    generator_get_return_default_null => {
        r#"<?php
function g(): Generator { yield 1; }
$g = g();
iterator_to_array($g);
echo $g->getReturn() === null ? 'null' : 'val';
"#,
        ["null"]
    };

    generator_yield_object_by_reference => {
        r#"<?php
function g(): Generator {
    $o = new stdClass();
    $o->n = 1;
    yield $o;
}
$obj = iterator_to_array(g())[0];
echo $obj->n;
"#,
        ["1"]
    };

    generator_in_class_method => {
        r#"<?php
class R {
    public function gen(): Generator { yield 'x'; }
}
echo iterator_to_array((new R())->gen())[0];
"#,
        ["x"]
    };

    generator_yield_from_string_chars => {
        r#"<?php
function g(): Generator { yield from 'ab'; }
echo implode('', iterator_to_array(g()));
"#,
        ["ab"]
    };

    generator_count_via_iterator => {
        r#"<?php
function g(): Generator { yield 1; yield 2; yield 3; }
echo iterator_count(g());
"#,
        ["3"]
    };

    generator_first_value_via_current => {
        r#"<?php
function g(): Generator { yield 'first'; yield 'second'; }
$g = g();
echo $g->current();
"#,
        ["first"]
    };

    generator_second_key_after_next => {
        r#"<?php
function g(): Generator { yield 'a'; yield 'b'; }
$g = g();
$g->next();
$g->next();
echo $g->key();
"#,
        ["1"]
    };

    generator_yield_list_destructure => {
        r#"<?php
function g(): Generator { yield [1, 2]; }
[$a, $b] = iterator_to_array(g())[0];
echo $a + $b;
"#,
        ["3"]
    };

    generator_closure_factory => {
        r#"<?php
$make = function (): Generator { yield 7; };
echo iterator_to_array($make())[0];
"#,
        ["7"]
    };

    generator_yield_from_associative_merge => {
        r#"<?php
function g(): Generator {
    yield from ['a' => 1];
    yield from ['b' => 2];
}
echo json_encode(iterator_to_array(g()));
"#,
        ["{\"a\":1,\"b\":2}"]
    };

    generator_empty_yields_nothing => {
        r#"<?php
function g(): Generator { return; }
echo count(iterator_to_array(g()));
"#,
        ["0"]
    };

    generator_yield_null_value => {
        r#"<?php
function g(): Generator { yield null; }
echo iterator_to_array(g())[0] === null ? 'null' : 'set';
"#,
        ["null"]
    };
}
