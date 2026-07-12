//! Generator `yield`, `yield from`, keys, and iterator protocol.

crate::php_cases! {
    generator_yield_three_values => {
        r#"<?php
function gen(): Generator { yield 1; yield 2; yield 3; }
echo implode(',', iterator_to_array(gen()));
"#,
        ["1,2,3"]
    };

    generator_yield_key_value_pairs => {
        r#"<?php
function gen(): Generator { yield 'a' => 1; yield 'b' => 2; }
echo json_encode(iterator_to_array(gen()));
"#,
        ["{\"a\":1,\"b\":2}"]
    };

    generator_current_before_advance => {
        r#"<?php
function gen(): Generator { yield 10; yield 20; }
$g = gen();
echo $g->current();
"#,
        ["10"]
    };

    generator_key_default_index => {
        r#"<?php
function gen(): Generator { yield 'x'; yield 'y'; }
$g = gen();
$g->next();
echo $g->key();
"#,
        ["1"]
    };

    generator_valid_false_after_exhausted => {
        r#"<?php
function gen(): Generator { yield 1; }
$g = gen();
iterator_to_array($g);
echo $g->valid() ? 'yes' : 'no';
"#,
        ["no"]
    };

    generator_return_value_via_get_return => {
        r#"<?php
function gen(): Generator { yield 1; return 99; }
$g = gen();
iterator_to_array($g);
echo $g->getReturn();
"#,
        ["99"]
    };

    generator_yield_from_delegates_array => {
        r#"<?php
function gen(): Generator { yield from [4, 5]; yield 6; }
echo implode(',', iterator_to_array(gen()));
"#,
        ["6,5"]
    };

    generator_yield_from_inner_generator => {
        r#"<?php
function inner(): Generator { yield 'a'; yield 'b'; }
function outer(): Generator { yield from inner(); yield 'c'; }
echo implode('', iterator_to_array(outer()));
"#,
        ["cb"]
    };

    generator_send_injects_value_after_yield => {
        r#"<?php
function gen(): Generator {
    $x = yield 'first';
    yield 'got:' . $x;
}
$g = gen();
$g->current();
$g->send('Z');
echo $g->current();
"#,
        ["got:Z"]
    };

    generator_throw_propagates_into_body => {
        r#"<?php
function gen(): Generator {
    try { yield 1; } catch (RuntimeException $e) { yield 'caught'; }
}
$g = gen();
$g->current();
$g->throw(new RuntimeException('x'));
echo $g->current();
"#,
        ["caught"]
    };

    generator_early_return_stops_further_yields => {
        r#"<?php
function gen(): Generator { yield 1; return; yield 2; }
echo implode(',', iterator_to_array(gen()));
"#,
        ["1"]
    };

    generator_in_foreach_loop => {
        r#"<?php
function gen(): Generator { yield 2; yield 4; }
$s = 0;
foreach (gen() as $v) { $s += $v; }
echo $s;
"#,
        ["6"]
    };

    generator_to_array_preserves_keys => {
        r#"<?php
function gen(): Generator { yield 'k' => 7; }
echo iterator_to_array(gen())['k'];
"#,
        ["7"]
    };

    generator_multiple_yield_from_sequence => {
        r#"<?php
function a(): Generator { yield 1; }
function b(): Generator { yield 2; }
function all(): Generator { yield from a(); yield from b(); }
echo implode('', iterator_to_array(all()));
"#,
        ["2"]
    };

    generator_get_return_throws_while_running => {
        r#"<?php
function gen(): Generator { yield 1; }
$g = gen();
$g->current();
try { $g->getReturn(); echo 'ok'; } catch (Exception) { echo 'err'; }
"#,
        ["err"]
    };
}
