//! `serialize` / `unserialize` roundtrips for arrays, objects, and scalars.

crate::php_cases! {
    serialize_unserialize_integer => {
        r#"<?php
echo unserialize(serialize(42));
"#,
        ["42"]
    };

    serialize_unserialize_string => {
        r#"<?php
echo unserialize(serialize('vybe'));
"#,
        ["vybe"]
    };

    serialize_unserialize_indexed_array => {
        r#"<?php
echo json_encode(unserialize(serialize([1, 2, 3])));
"#,
        ["[1,2,3]"]
    };

    serialize_unserialize_assoc_array => {
        r#"<?php
$a = unserialize(serialize(['k' => 'v']));
echo $a['k'];
"#,
        ["v"]
    };

    serialize_unserialize_null => {
        r#"<?php
echo unserialize(serialize(null)) === null ? 'null' : 'val';
"#,
        ["null"]
    };

    serialize_unserialize_boolean_true => {
        r#"<?php
echo unserialize(serialize(true)) ? 'true' : 'false';
"#,
        ["true"]
    };

    serialize_unserialize_stdclass_property => {
        r#"<?php
$o = new stdClass();
$o->x = 9;
echo unserialize(serialize($o))->x;
"#,
        ["9"]
    };

    serialize_format_starts_with_type_marker => {
        r#"<?php
echo str_starts_with(serialize(1), 'i:') ? 'int' : 'other';
"#,
        ["int"]
    };

    serialize_object_format_marker => {
        r#"<?php
echo str_starts_with(serialize(new stdClass()), 'O:') ? 'obj' : 'other';
"#,
        ["obj"]
    };

    unserialize_false_on_garbage => {
        r#"<?php
echo unserialize('not-serialized') === false ? 'false' : 'ok';
"#,
        ["false"]
    };

    serialize_preserves_nested_array_shape => {
        r#"<?php
$d = unserialize(serialize(['a' => ['b' => 2]]));
echo $d['a']['b'];
"#,
        ["2"]
    };

    serialize_float_value => {
        r#"<?php
echo unserialize(serialize(1.5));
"#,
        ["1.5"]
    };

    serialize_empty_array => {
        r#"<?php
echo json_encode(unserialize(serialize([])));
"#,
        ["[]"]
    };

    serialize_object_with_private_like_dynamic_prop => {
        r#"<?php
class Bag { public int $n = 5; }
echo unserialize(serialize(new Bag()))->n;
"#,
        ["5"]
    };

    serialize_recursion_not_in_simple_tree => {
        r#"<?php
$a = [1];
echo unserialize(serialize($a)) === $a ? 'same' : 'diff';
"#,
        ["same"]
    };
}
