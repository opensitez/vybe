//! List/array destructuring assignment — real PHP 8.4 semantics.
//!
//! Verified against the `php` 8.4 CLI: destructuring from a scalar (int, float,
//! bool, null, string, resource) does NOT throw — every target is assigned
//! `null`. Destructuring from an object or a Generator throws base `\Error`
//! ("Cannot use object of type X as array"), NOT `TypeError`. A typed-property
//! target that receives an incompatible value throws `TypeError`. Missing keys
//! / short arrays leave `null`. Earlier revisions asserted `TypeError` for the
//! scalar sources, which real PHP does not do.

crate::php_cases! {
    list_destructure_from_integer_assigns_null => {
        r#"<?php
[$a] = 42;
echo $a === null ? 'null' : 'set';
"#,
        ["null"]
    };

    list_destructure_from_null_assigns_null => {
        r#"<?php
[$a, $b] = null;
echo ($a === null && $b === null) ? 'nulls' : 'set';
"#,
        ["nulls"]
    };

    list_destructure_from_object_throws_error => {
        r#"<?php
try { [$x] = new stdClass(); echo 'ok'; }
catch (\Error $e) { echo 'obj-src'; }
"#,
        ["obj-src"]
    };

    list_destructure_from_false_assigns_null => {
        r#"<?php
[$t] = false;
echo $t === null ? 'null' : 'set';
"#,
        ["null"]
    };

    list_destructure_from_float_assigns_null => {
        r#"<?php
[$f] = 1.5;
echo $f === null ? 'null' : 'set';
"#,
        ["null"]
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

    nested_list_destructure_from_scalar_inner_assigns_null => {
        r#"<?php
[[$a, $b]] = [1, 2];
echo ($a === null && $b === null) ? 'nulls' : 'set';
"#,
        ["nulls"]
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

    list_destructure_on_resource_assigns_null => {
        r#"<?php
$fp = fopen('php://memory', 'r+');
[$r] = $fp;
echo $r === null ? 'null' : 'set';
fclose($fp);
"#,
        ["null"]
    };

    keyed_destructure_on_scalar_assigns_null => {
        r#"<?php
['k' => $v] = 'str';
echo $v === null ? 'null' : 'set';
"#,
        ["null"]
    };

    list_destructure_nested_keyed_inner_missing => {
        r#"<?php
[['id' => $id], $tail] = [['name' => 'x'], 1];
echo $id === null ? 'no-id' : $id;
"#,
        ["no-id"]
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

    list_destructure_from_generator_throws_error => {
        r#"<?php
function g(): Generator { yield 1; yield 2; }
try { [$a, $b] = g(); echo 'ok'; }
catch (\Error $e) { echo 'gen-src'; }
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

    list_destructure_string_too_short_assigns_null => {
        r#"<?php
[$a, $b] = 'x';
echo ($a === null && $b === null) ? 'nulls' : 'set';
"#,
        ["nulls"]
    };
}
