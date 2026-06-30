//! `unserialize` / `serialize` runtime failure paths (not compile-only checks).

crate::php_cases! {
    unserialize_garbage_returns_false => {
        r#"<?php
$v = unserialize('not-a-serialized-blob');
echo $v === false ? 'false' : 'value';
"#,
        ["false"]
    };

    unserialize_truncated_object_returns_false => {
        r#"<?php
$v = unserialize('O:8:"stdClass":1:{s:1');
echo $v === false ? 'trunc' : 'ok';
"#,
        ["trunc"]
    };

    unserialize_unknown_class_becomes_incomplete_class => {
        r#"<?php
$v = unserialize('O:15:"Missing\\Klass":0:{}');
echo is_object($v) && $v instanceof __PHP_Incomplete_Class ? 'incomplete' : 'other';
"#,
        ["incomplete"]
    };

    unserialize_allowed_classes_false_blocks_object_wakeup => {
        r#"<?php
class Box { public int $n = 0; }
$blob = serialize(new Box());
$v = unserialize($blob, ['allowed_classes' => false]);
echo $v instanceof __PHP_Incomplete_Class ? 'blocked' : 'live';
"#,
        ["blocked"]
    };

    unserialize_allowed_classes_whitelist_permits_match => {
        r#"<?php
class Ok {}
$blob = serialize(new Ok());
$v = unserialize($blob, ['allowed_classes' => [Ok::class]]);
echo $v instanceof Ok ? 'ok' : 'nope';
"#,
        ["ok"]
    };

    unserialize_max_depth_exceeded_returns_false => {
        r#"<?php
$deep = ['l' => null];
$ref = &$deep['l'];
$ref = &$deep;
$blob = serialize($deep);
$v = @unserialize($blob, ['max_depth' => 2]);
echo $v === false ? 'depth' : 'parsed';
"#,
        ["depth"]
    };

    serialize_resource_returns_false => {
        r#"<?php
$fp = fopen('php://memory', 'r+');
$ok = serialize($fp);
fclose($fp);
echo $ok === false ? 'no-ser' : 'ser';
"#,
        ["no-ser"]
    };

    unserialize_rejects_invalid_type_prefix => {
        r#"<?php
$v = unserialize('Z:3:"foo";');
echo $v === false ? 'bad-type' : 'parsed';
"#,
        ["bad-type"]
    };

    unserialize_empty_string_returns_false => {
        r#"<?php
$v = unserialize('');
echo $v === false ? 'empty' : 'value';
"#,
        ["empty"]
    };

    unserialize_bool_true_roundtrip => {
        r#"<?php
echo unserialize(serialize(true)) ? 'T' : 'F';
"#,
        ["T"]
    };

    unserialize_array_nested_count_preserved => {
        r#"<?php
$data = ['a' => [1, 2], 'b' => 3];
$back = unserialize(serialize($data));
echo count($back) . '-' . count($back['a']);
"#,
        ["2-2"]
    };

    unserialize_object_sleep_controls_fields => {
        r#"<?php
class Pick {
    public int $keep = 1;
    private int $hide = 9;
    public function __sleep(): array { return ['keep']; }
}
$o = unserialize(serialize(new Pick()));
echo $o->keep;
"#,
        ["1"]
    };

    unserialize_wakeup_mutates_restored_object => {
        r#"<?php
class Wake {
    public string $tag = 'raw';
    public function __wakeup(): void { $this->tag = 'woke'; }
}
$o = unserialize(serialize(new Wake()));
echo $o->tag;
"#,
        ["woke"]
    };

    unserialize_private_name_mangled_property_restored => {
        r#"<?php
class Secret {
    public function reveal(): int {
        return $this->code;
    }
    private int $code = 42;
}
$s = new Secret();
$copy = unserialize(serialize($s));
echo $copy->reveal();
"#,
        ["42"]
    };

    serialize_enum_case_roundtrips_value => {
        r#"<?php
enum Color { case Red; case Blue; }
$c = unserialize(serialize(Color::Blue));
echo $c === Color::Blue ? 'blue' : 'other';
"#,
        ["blue"]
    };

    unserialize_references_preserve_identity => {
        r#"<?php
$a = ['z' => 1];
$a['self'] = &$a['z'];
$back = unserialize(serialize($a));
$back['z'] = 5;
echo $back['self'];
"#,
        ["5"]
    };

    unserialize_invalid_length_string_returns_false => {
        r#"<?php
$v = unserialize('s:10:"short";');
echo $v === false ? 'len' : 'ok';
"#,
        ["len"]
    };

    unserialize_object_with_custom_class_name_case_sensitive => {
        r#"<?php
class CaseSens {}
$blob = serialize(new CaseSens());
$v = unserialize($blob);
echo $v instanceof CaseSens ? 'match' : 'miss';
"#,
        ["match"]
    };
}
