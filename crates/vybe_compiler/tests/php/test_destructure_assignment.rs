//! List/array destructuring assignment failures — TypeErrors and null slots only.

crate::php_cases! {
    list_destructure_from_integer_throws_type_error => {
        r#"<?php
try { [$a] = 42; echo 'ok'; }
catch (TypeError $e) { echo 'int-src'; }
"#,
        ["int-src"]
    };

    list_destructure_from_null_throws_type_error => {
        r#"<?php
try { [$a, $b] = null; echo 'ok'; }
catch (TypeError $e) { echo 'null-src'; }
"#,
        ["null-src"]
    };

    list_destructure_from_object_throws_type_error => {
        r#"<?php
try { [$x] = new stdClass(); echo 'ok'; }
catch (TypeError $e) { echo 'obj-src'; }
"#,
        ["obj-src"]
    };

    list_destructure_from_false_throws_type_error => {
        r#"<?php
try { [$t] = false; echo 'ok'; }
catch (TypeError $e) { echo 'false-src'; }
"#,
        ["false-src"]
    };

    list_destructure_from_float_throws_type_error => {
        r#"<?php
try { [$f] = 1.5; echo 'ok'; }
catch (TypeError $e) { echo 'float-src'; }
"#,
        ["float-src"]
    };

    list_destructure_fewer_source_elements_leaves_null => {
        r#"<?php
[$a, $b, $c] = [1];
echo ($b === null && $c === null) ? 'nulls' : 'filled';
"#,
        ["nulls"]
    };

    keyed_destructure_missing_key_yields_null => {
        r#"<?php
['need' => $v] = ['have' => 1];
echo $v === null ? 'missing' : 'found';
"#,
        ["missing"]
    };

    nested_list_destructure_requires_array_inner => {
        r#"<?php
try { [[$a, $b]] = [1, 2]; echo 'ok'; }
catch (TypeError $e) { echo 'nested'; }
"#,
        ["nested"]
    };

    list_destructure_empty_array_leaves_all_null => {
        r#"<?php
[$p, $q] = [];
echo ($p === null && $q === null) ? 'both-null' : 'set';
"#,
        ["both-null"]
    };

    list_destructure_spread_must_be_last_parse_error => {
        r#"<?php
try {
    eval('[$a, ...$b, $c] = [1,2,3];');
    echo 'ok';
} catch (ParseError $e) {
    echo 'parse';
}
"#,
        ["parse"]
    };

    list_destructure_on_resource_throws_type_error => {
        r#"<?php
$fp = fopen('php://memory', 'r+');
try { [$r] = $fp; echo 'ok'; }
catch (TypeError $e) { echo 'res-src'; }
finally { fclose($fp); }
"#,
        ["res-src"]
    };

    keyed_destructure_on_scalar_throws_type_error => {
        r#"<?php
try { ['k' => $v] = 'str'; echo 'ok'; }
catch (TypeError $e) { echo 'key-scalar'; }
"#,
        ["key-scalar"]
    };

    list_destructure_nested_keyed_inner_missing => {
        r#"<?php
[['id' => $id], $tail] = [['name' => 'x'], 1];
echo $id === null ? 'no-id' : $id;
"#,
        ["no-id"]
    };

    list_destructure_spread_from_non_array_throws => {
        r#"<?php
try { [$head, ...$rest] = 1; echo 'ok'; }
catch (TypeError $e) { echo 'spread-int'; }
"#,
        ["spread-int"]
    };

    list_destructure_typed_property_mismatch_throws => {
        r#"<?php
class Holder { public int $n; }
$h = new Holder();
try { [$h->n] = ['not-int']; echo 'ok'; }
catch (TypeError $e) { echo 'typed'; }
"#,
        ["typed"]
    };

    list_destructure_readonly_property_outside_constructor => {
        r#"<?php
class Ro { public function __construct(public readonly int $v) {} }
$o = new Ro(1);
try { [$o->v] = [2]; echo 'ok'; }
catch (Error $e) { echo 'readonly'; }
"#,
        ["readonly"]
    };

    list_destructure_from_generator_not_directly_valid => {
        r#"<?php
function g(): Generator { yield 1; yield 2; }
try { [$a, $b] = g(); echo 'ok'; }
catch (TypeError $e) { echo 'gen-src'; }
"#,
        ["gen-src"]
    };

    list_destructure_extra_outer_slots_stay_null => {
        r#"<?php
[$one, $two, $three] = [9];
echo ($one === 9 && $two === null && $three === null) ? 'shape' : 'wrong';
"#,
        ["shape"]
    };

    keyed_destructure_null_value_not_same_as_missing_key => {
        r#"<?php
['z' => $z] = ['z' => null];
echo $z === null ? 'null-val' : 'set';
"#,
        ["null-val"]
    };

    list_destructure_string_too_short_for_two_slots => {
        r#"<?php
try { [$a, $b] = 'x'; echo 'ok'; }
catch (ValueError $e) { echo 'short-str'; }
"#,
        ["short-str"]
    };
}

