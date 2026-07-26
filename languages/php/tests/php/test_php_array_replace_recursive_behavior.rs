use super::helpers::run_prints;

#[test]
fn test_array_replace_recursive_nested_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$base = ['db' => ['host' => 'localhost', 'port' => 3306]];
$custom = ['db' => ['host' => '127.0.0.1', 'user' => 'root']];
$result = array_replace_recursive($base, $custom);
echo $result['db']['host'] . ':' . $result['db']['port'] . ':' . $result['db']['user'], "\n";
"#
        ),
        vec!["127.0.0.1:3306:root"]
    );
}

#[test]
fn test_array_replace_recursive_scalar_overwrites_array() {
    assert_eq!(
        run_prints(
            r#"<?php
$base = ['config' => ['a' => 1, 'b' => 2]];
$custom = ['config' => 'disabled'];
$result = array_replace_recursive($base, $custom);
echo is_string($result['config']) ? $result['config'] : 'array', "\n";
"#
        ),
        vec!["disabled"]
    );
}

#[test]
fn test_array_replace_recursive_preserves_new_nested_key() {
    assert_eq!(
        run_prints(
            r#"<?php
$base = ['app' => ['env' => 'prod', 'debug' => false]];
$override = ['app' => ['locale' => 'en', 'debug' => true]];
$result = array_replace_recursive($base, $override);
echo $result['app']['env'] . '|' . ($result['app']['debug'] ? '1' : '0') . '|' . $result['app']['locale'];
"#
        ),
        vec!["prod|1|en"]
    );
}

#[test]
fn test_array_replace_recursive_multiple_sources_override_order() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['x' => 1, 'nested' => ['a' => 1, 'b' => 2]];
$b = ['y' => 9, 'nested' => ['a' => 2]];
$c = ['nested' => ['b' => 3], 'x' => 4];
$result = array_replace_recursive($a, $b, $c);
echo $result['x'] . '|' . $result['y'] . '|' . $result['nested']['a'] . '|' . $result['nested']['b'];
"#
        ),
        vec!["4|9|2|3"]
    );
}

#[test]
fn test_array_replace_recursive_numeric_keys_are_replaced() {
    assert_eq!(
        run_prints(
            r#"<?php
$base = [0 => 'zero', 1 => 'one', 2 => [ 'inner' => 'base' ]];
$patch = [1 => 'uno', 2 => ['inner' => 'patched', 'added' => 'yes']];
$result = array_replace_recursive($base, $patch);
echo $result[0] . '|' . $result[1] . '|' . $result[2]['inner'] . '|' . $result[2]['added'];
"#
        ),
        vec!["zero|uno|patched|yes"]
    );
}

#[test]
fn test_array_replace_recursive_null_value_replaces_nested() {
    assert_eq!(
        run_prints(
            r#"<?php
$base = ['cfg' => ['a' => 1, 'b' => 2], 'flag' => true];
$patch = ['cfg' => null];
$result = array_replace_recursive($base, $patch);
echo is_null($result['cfg']) ? 'null' : 'array';
echo '|' . ($result['flag'] ? '1' : '0');
"#
        ),
        vec!["null|1"]
    );
}

#[test]
fn test_array_replace_recursive_scalar_to_array_is_not_recursive() {
    assert_eq!(
        run_prints(
            r#"<?php
$base = ['cfg' => 'scalar'];
$patch = ['cfg' => ['next' => 9]];
$result = array_replace_recursive($base, $patch);
echo is_array($result['cfg']) ? $result['cfg']['next'] : 'scalar';
"#
        ),
        vec!["9"]
    );
}

#[test]
fn test_array_replace_recursive_associative_missing_and_empty() {
    assert_eq!(
        run_prints(
            r#"<?php
$base = ['a' => ['x' => 1], 'b' => ['y' => 2]];
$patch = ['a' => ['z' => 3], 'c' => ['w' => 4]];
$result = array_replace_recursive($base, $patch);
echo $result['a']['x'] . '|' . $result['a']['z'] . '|' . $result['b']['y'] . '|' . $result['c']['w'];
"#
        ),
        vec!["1|3|2|4"]
    );
}
