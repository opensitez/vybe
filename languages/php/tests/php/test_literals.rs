use super::helpers::run_prints;

fn assert_code(src: &str, expected: Vec<&str>) {
    let result = run_prints(src);
    assert_eq!(
        result,
        expected.into_iter().map(|s| s.to_string()).collect::<Vec<_>>()
    );
}

#[test]
fn test_php_integer_literals_print() {
    assert_code(
        "<?php\necho 127;\necho '\\n';\necho 0x2A;\necho '\\n';\necho 0o77;\necho '\\n';\necho 0b1010;\necho '\\n';\necho 1_000_000;",
        vec!["127", "42", "63", "10", "1000000"],
    );
}

#[test]
fn test_php_float_literals_print() {
    assert_code(
        "<?php\necho 3.5;\necho '\\n';\necho 1e2;\necho '\\n';\necho 1.2e-3;\necho '\\n';\necho (1 + 2.5);",
        vec!["3.5", "100", "0.0012", "3.5"],
    );
}

#[test]
fn test_php_boolean_and_null_literals_print() {
    assert_code(
        "<?php\nif (true) { echo 't'; } else { echo 'f'; }\necho '\\n';\nif (false) { echo 'bad'; } else { echo 'good'; }\necho '\\n';\nvar_dump(null);",
        vec!["t", "good", "NULL"],
    );
}

#[test]
fn test_php_string_literals_and_escape_sequences() {
    assert_code(
        "<?php\n$s = 'hi\\nthere';\necho $s;\necho '\\n';\n$name = 'B';\n$t = \"A{$name}\";\necho $t;\necho '\\n';\necho \"abc\\nxyz\";",
        vec!["hi\\nthere", "AB", "abc\\nxyz"],
    );
}

#[test]
fn test_php_heredoc_and_nowdoc_literals() {
    assert_code(
        "<?php\n$h = <<<TXT\nfirst\nsecond\nTXT;\necho str_replace('\\n', '|', $h);\necho '\\n';\n$n = <<<'NOWDOC'\nRAW\nNOWDOC;\necho $n;",
        vec!["first|second|", "RAW"],
    );
}

#[test]
fn test_php_array_literal_variants() {
    assert_code(
        "<?php\n$a = [1, 2, 3];\n$b = ['x' => 10, 'y' => 20];\n$c = [0 => 'zero', 2 => 'two', 'x' => 'ex'];\n\necho count($a);\necho '\\n';\necho $b['x'];\necho '\\n';\n$merged = $a + [2 => 5, 3 => 6];\necho json_encode($merged);\n",
        vec!["3", "10", "[1,2,3,6]"],
    );
}

#[test]
fn test_php_string_offset_access_and_assignment() {
    assert_code(
        "<?php\n$s = 'abc';\necho $s[0];\necho '\\n';\necho $s[1];\n$s[1] = 'Z';\necho '\\n';\necho $s;",
        vec!["a", "b", "aZc"],
    );
}

#[test]
fn test_php_boolean_logic_literals() {
    assert_code(
        "<?php\necho (true ? 'T' : 'F');\necho '\\n';\necho (false ? 'T' : 'F');\necho '\\n';\nvar_export(NULL === null);\n",
        vec!["T", "F", "true"],
    );
}

#[test]
fn test_php_numeric_string_and_float_literals() {
    assert_code(
        "<?php\necho 0b111;\necho '\\n';\necho 012;\necho '\\n';\necho 0x1a;\necho '\\n';\necho 1e-3;\necho '\\n';\necho 3.25e2;\n",
        vec!["7", "10", "26", "0.001", "325"],
    );
}

#[test]
fn test_php_constant_null_false_true_literals() {
    assert_code(
        "<?php\nvar_export(NULL);\necho '\\n';\nvar_export(TRUE);\necho '\\n';\nvar_export(FALSE);\necho '\\n';\nvar_export(!FALSE);\n",
        vec!["NULL", "true", "false", "true"],
    );
}

#[test]
fn test_php_array_shape_and_keyed_literals() {
    assert_code(
        "<?php\n$a = [0 => 'zero', 'x' => 1, 2 => 'two'];\n$b = [\n    'left' => ['k' => 1],\n    'right' => ['v' => 2],\n];\narray_push($a, 'tail');\necho $a[0];\necho '\\n';\necho $a['x'];\necho '\\n';\necho $b['left']['k'];\necho '\\n';\necho $a[3];\n",
        vec!["zero", "1", "1", "tail"],
    );
}

#[test]
fn test_php_string_interpolation_complex() {
    assert_code(
        "<?php\n$name = 'Ada';\n$age = 30;\necho \"Name: {$name}, Age: {$age}\";\necho '\\n';\necho \"{$name}\" . ' has ' . $age . ' years';\n",
        vec!["Name: Ada, Age: 30", "Ada has 30 years"],
    );
}

#[test]
fn test_php_binary_and_hexdump_literals_like_unicode() {
    assert_code(
        "<?php\necho \"\\x48\\x69\";\necho '|';\necho \"\u{2665}\";\n",
        vec!["Hi|♥"],
    );
}

#[test]
fn test_php_single_quoted_literal_passthrough() {
    assert_code(
        "<?php\necho 'a\\\\nb';\necho '\\n';\necho '\\n';\necho 'line\\\\' . \"\\n\";\n",
        vec!["a\\nb", "", "line\\"],
    );
}

#[test]
fn test_php_boolean_cast_literals_and_is_numeric() {
    assert_code(
        "<?php\necho (int) true;\necho '|';\necho (int) false;\necho '|';\necho is_numeric('1_000') ? 'n' : 'v';\necho '|';\necho is_numeric('1000') ? 'n' : 'v';\n",
        vec!["1|0|v|n"],
    );
}

#[test]
fn test_php_float_precision_and_rounding_literals() {
    assert_code(
        "<?php\necho number_format(1.5 + 2.5, 0);\necho '|';\necho sprintf('%.1f', 3.14159);\n",
        vec!["4|3.1"],
    );
}

#[test]
fn test_php_nested_heredoc_literals() {
    assert_code(
        "<?php\n$xml = <<<XML\n<root>\n  <node/>\n</root>\nXML;\necho str_replace(\"\\n\", \"\", $xml);\n",
        vec!["<root>  <node/></root>"],
    );
}

#[test]
fn test_php_empty_string_and_nullish_coalescing_with_literals() {
    assert_code(
        "<?php\necho strlen('');\necho '|';\necho (null ?? 'default');\necho '|';\necho ('' ?: 'fallback');\n",
        vec!["0| |fallback"],
    );
}

#[test]
fn test_php_numeric_literal_mix_runtime() {
    assert_code(
        "<?php\necho 0b10 + 0o10 + 0x10 + 1_000;\necho '|';\necho sprintf('%.1f', 1_000.5 + 2_000.25);\necho '|';\necho (int)'1_000';\n",
        vec!["1034", "3000.8", "1"],
    );
}

#[test]
fn test_php_magic_constants_in_class_scope() {
    assert_code(
        r#"<?php
namespace NS;
class Sample {
    public function identity(): string {
        return __CLASS__ . ':' . __FUNCTION__ . ':' . __METHOD__ . ':' . __NAMESPACE__ . ':' . __TRAIT__;
    }
}
$obj = new Sample();
echo $obj->identity();
"#,
        vec!["NS\\Sample:identity:Sample::identity:NS:"],
    );
}

#[test]
fn test_php_heredoc_and_nowdoc_indentation() {
    assert_code(
        r#"<?php
$heredoc = <<<TXT
  indented
    line
TXT;
$nowdoc = <<<'RAW'
  raw
    stop
RAW;
echo $heredoc;
echo '|';
echo str_replace("\n", ",", $nowdoc);
"#,
        vec!["  indented\n    line|  raw,    stop"],
    );
}

#[test]
fn test_php_array_literals_with_duplicate_keys_and_trailing_comma() {
    assert_code(
        "<?php\n$values = ['a' => 1, 'b' => 2, 'a' => 3];\necho $values['a'];\necho '|';\necho $values['b'];\n",
        vec!["3|2"],
    );
}

#[test]
fn test_php_integer_literal_prefixes_and_underscores() {
    assert_code(
        "<?php\necho 0b1001;\necho '|';\necho 0o17;\necho '|';\necho 0x1f;\necho '|';\necho 1_024_000;\n",
        vec!["9|15|31|1024000"],
    );
}

#[test]
fn test_php_float_and_nan_in_literals() {
    assert_code(
        "<?php\necho round(1.25 + 0.5, 1);\necho '|';\necho is_float(1.0 / 2);\necho '|';\necho (NAN === NAN) ? 'nan' : 'not';\n",
        vec!["1.8|1|not"],
    );
}

#[test]
fn test_php_complex_nested_array_shape_with_mixed_literals() {
    assert_code(
        "<?php\n$data = [\n    'id' => 1,\n    'meta' => ['name' => 'A', 'active' => true],\n    5 => 'num',\n];\necho $data['id'];\necho '|';\necho $data['meta']['name'];\necho '|';\necho $data[5];\necho '|';\necho $data['meta']['active'] ? 'on' : 'off';\n",
        vec!["1|A|num|on"],
    );
}

#[test]
fn test_php_octal_string_index_vs_key_cast_literals() {
    assert_code(
        "<?php\n$a = ['01' => 'string', 1 => 'one', 01 => 'zero'];\necho array_key_exists('1', $a) ? 'has1' : 'no1';\necho '|';\necho $a['1'];\necho '|';\necho $a[1];\n",
        vec!["has1|zero|zero"],
    );
}

#[test]
fn test_php_magic_constant_literals_in_function_scope() {
    assert_code(
        "<?php\nfunction marker() {\n    return __LINE__ . ':' . __FILE__ . ':' . __FUNCTION__;\n}\necho strpos(marker(), ':') > 0 ? 'yes' : 'no';\n",
        vec!["yes"],
    );
}

#[test]
fn test_php_octal_and_hex_string_key_edge_cases() {
    assert_code(
        "<?php\n$a = [0o10 => 'octal-key', 010 => 'legacy-octal', 0x10 => 'hex'];\necho $a[8];\necho '|';\necho $a[16];\necho '|';\necho $a['10'];\n",
        vec!["octal-key|hex|octal-key"],
    );
}

#[test]
fn test_php_negative_numeric_literals_and_operator_interaction() {
    assert_code(
        "<?php\necho -123;\necho '|';\necho -0b10;\necho '|';\necho -0xF;\necho '|';\necho -(0b10 + 0x2) * 2;\n",
        vec!["-123|-2|-15|-8"],
    );
}

#[test]
fn test_php_float_scientific_variants() {
    assert_code(
        "<?php\necho 1e3;\necho '|';\necho 1.2E2;\necho '|';\necho 5e-1;\necho '|';\necho 1_000.5;\n",
        vec!["1000|120|0.5|1000.5"],
    );
}

#[test]
fn test_php_string_literals_with_hex_and_unicode_escapes() {
    assert_code(
        "<?php\necho \"A\\x42C\";\necho '|';\necho \"\\u{2764}\";\necho '|';\necho \"\u{1F600}\";\n",
        vec!["ABC|❤|😀"],
    );
}

#[test]
fn test_php_multi_line_nowdoc_with_trailing_spacing() {
    assert_code(
        r#"<?php
$payload = <<<'JSON'
{
  "ok": true
}
JSON;
echo str_replace("\n", ";", $payload);
"#,
        vec!["{;  \"ok\": true;};"],
    );
}
