use super::helpers::{compile_ok, run_prints};

// ── String operations ───────────────────────────────────────
#[test]
fn concat_dot() {
    compile_ok("<?php $x = 'hello' . ' ' . 'world'; echo $x;");
}
#[test]
fn concat_assign() {
    compile_ok("<?php $x = 'hello'; $x .= ' world'; echo $x;");
}
#[test]
fn string_repeat() {
    compile_ok("<?php echo str_repeat('ab', 3);");
}
#[test]
fn string_length() {
    compile_ok("<?php echo strlen('hello');");
}
#[test]
fn string_case() {
    compile_ok("<?php echo strtolower('HELLO'); echo strtoupper('hello');");
}
#[test]
fn string_trim() {
    compile_ok("<?php echo trim('  hi  '); echo ltrim('  hi'); echo rtrim('hi  ');");
}
#[test]
fn string_substr() {
    compile_ok("<?php echo substr('hello world', 6); echo substr('hello', 0, 3);");
}
#[test]
fn string_substr_dynamic_receiver_runtime() {
    assert_eq!(
        run_prints(
            "<?php $_SERVER = ['HTTP_ACCEPT_LANGUAGE' => 'abcdef']; echo substr($_SERVER['HTTP_ACCEPT_LANGUAGE'], 0, 2);"
        ),
        vec!["ab".to_string()]
    );
}
#[test]
fn string_search() {
    compile_ok("<?php echo strpos('hello', 'lo'); echo str_contains('hello', 'ell');");
}
#[test]
fn string_collation_compare_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php echo strcoll('A', 'B'); echo "\n"; echo strcoll('B', 'A'); echo "\n"; echo strcoll('A', 'A'); "#
        ),
        vec!["-1".to_string(), "1".to_string(), "0".to_string()]
    );
}
#[test]
fn string_replace() {
    compile_ok("<?php echo str_replace('world', 'PHP', 'hello world');");
}
#[test]
fn string_split() {
    compile_ok("<?php $parts = explode(',', 'a,b,c'); echo $parts[0];");
}
#[test]
fn string_join() {
    compile_ok("<?php echo implode('-', ['a', 'b', 'c']);");
}
#[test]
fn string_starts_ends() {
    compile_ok("<?php echo str_starts_with('hello', 'he'); echo str_ends_with('hello', 'lo');");
}
#[test]
fn string_pad() {
    compile_ok("<?php echo str_pad('42', 5, '0');");
}
#[test]
fn string_ucfirst() {
    compile_ok("<?php echo ucfirst('hello');");
}
#[test]
fn string_lcfirst() {
    compile_ok("<?php echo lcfirst('Hello');");
}

// ── String interpolation ────────────────────────────────────
#[test]
fn interp_var() {
    compile_ok(r#"<?php $name = "World"; echo "Hello $name!";"#);
}
#[test]
fn interp_curly() {
    compile_ok(r#"<?php $x = "test"; echo "val: {$x}";"#);
}
#[test]
fn interp_array() {
    compile_ok(r#"<?php $a = ['hi']; echo "val: {$a[0]}";"#);
}
#[test]
fn interp_prop() {
    compile_ok(r#"<?php $o = new stdClass(); echo "val: $o->name";"#);
}
#[test]
fn interp_escape() {
    compile_ok(r#"<?php echo "price: \$5";"#);
}
#[test]
fn interp_special_chars() {
    compile_ok(r#"<?php echo "line1\nline2\ttab";"#);
}
#[test]
fn no_interp_single_quote() {
    compile_ok("<?php echo 'no $interpolation here';");
}

// ── Heredoc / Nowdoc ────────────────────────────────────────
#[test]
fn heredoc() {
    compile_ok("<?php $x = <<<EOT\nHello World\nEOT;\necho $x;");
}
#[test]
fn nowdoc() {
    compile_ok("<?php $x = <<<'EOT'\nNo $interpolation\nEOT;\necho $x;");
}

// ── Encoding functions ──────────────────────────────────────
#[test]
fn htmlspecialchars() {
    compile_ok("<?php echo htmlspecialchars('<b>hi</b>');");
}
#[test]
fn urlencode_decode() {
    compile_ok("<?php $e = urlencode('hello world'); echo urldecode($e);");
}
#[test]
fn base64() {
    compile_ok("<?php $e = base64_encode('hello'); echo base64_decode($e);");
}
#[test]
fn json_roundtrip() {
    compile_ok("<?php $j = json_encode(['a'=>1]); $d = json_decode($j);");
}
#[test]
fn nl2br() {
    compile_ok("<?php echo nl2br(\"line1\\nline2\");");
}
#[test]
fn sprintf_format() {
    compile_ok("<?php echo sprintf('Hello %s, age %d', 'John', 30);");
}

#[test]
fn implode_and_explode_runtime() {
    assert_eq!(
        run_prints("<?php $parts = explode('|', 'a|b|c|d', 3); echo implode('-', $parts);"),
        vec!["a-b-c|d".to_string()]
    );
}

#[test]
fn string_trim_character_mask_runtime() {
    assert_eq!(
        run_prints("<?php echo trim('__hello__', '_');"),
        vec!["hello".to_string()]
    );
}

#[test]
fn string_substr_negative_offset_runtime() {
    assert_eq!(
        run_prints("<?php echo substr('abcdef', -3); echo \"\\n\"; echo substr('abcdef', -3, 1);"),
        vec!["def".to_string(), "d".to_string()]
    );
}

#[test]
fn string_find_and_match_runtime() {
    assert_eq!(
        run_prints(
            "<?php $s = 'hello world'; echo str_contains($s, 'world') ? '1' : '0'; echo \"\\n\"; echo str_starts_with($s, 'he') ? '1' : '0'; echo \"\\n\"; echo str_ends_with($s, 'ld') ? '1' : '0';"
        ),
        vec!["1".to_string(), "1".to_string(), "1".to_string()]
    );
}

#[test]
fn string_regex_replace_runtime() {
    assert_eq!(
        run_prints("<?php echo preg_replace('/(\\w+)@(\\w+)/', '$2@$1', 'a@b');"),
        vec!["b@a".to_string()]
    );
}

#[test]
fn string_split_limit_runtime() {
    assert_eq!(
        run_prints("<?php echo substr_count('aaa', 'a'); echo \"\\n\"; echo str_word_count('one two  three');"),
        vec!["3".to_string(), "3".to_string()]
    );
}

#[test]
fn string_escaping_quotes_runtime() {
    assert_eq!(
        run_prints("<?php $s = \"a\\t b\\n c\"; echo json_encode($s); echo \"\\n\"; echo addcslashes('a b', ' a');"),
        vec!["\"a\\t b\\n c\"".to_string(), "a\\ b".to_string()]
    );
}

#[test]
fn string_pointer_runtime() {
    assert_eq!(
        run_prints("<?php $s = 'hello'; echo strlen($s); echo \"\\n\"; $c = strlen($s) + 0; echo $c;"),
        vec!["5".to_string(), "5".to_string()]
    );
}

#[test]
fn string_replace_callback_runtime() {
    assert_eq!(
        run_prints("<?php echo preg_replace_callback('/(\\d+)/', fn($m) => strval((int)$m[1] + 1), 'x1y2');"),
        vec!["x2y3".to_string()]
    );
}

#[test]
fn string_explode_preserves_empty_fields_runtime() {
    assert_eq!(
        run_prints(
            "<?php
$parts = explode(',', 'a,,b,');
echo count($parts);
echo '|';
echo $parts[1] === '' ? 'empty' : 'filled';
echo '|';
echo isset($parts[3]) ? ($parts[3] === '' ? 'tail-empty' : 'tail-fill') : 'tail-miss';
"
        ),
        vec!["4|empty|tail-empty".to_string()]
    );
}

#[test]
fn string_implode_with_nested_arrays_runtime() {
    assert_eq!(
        run_prints(
            "<?php
$parts = [1, 'x', 3];
echo implode(':', $parts);
"
        ),
        vec!["1:x:3".to_string()]
    );
}

#[test]
fn string_trim_custom_char_mask_runtime() {
    assert_eq!(
        run_prints(
            "<?php
echo trim('__a-b-c__', '_');
"
        ),
        vec!["a-b-c".to_string()]
    );
}

#[test]
fn string_repeat_negative_guard_runtime() {
    assert_eq!(
        run_prints(
            "<?php
echo str_repeat('x', 0);
echo '|';
echo str_repeat('y', 3);
"
        ),
        vec!["|yyy".to_string()]
    );
}

#[test]
fn string_to_upper_lower_runtime() {
    assert_eq!(
        run_prints(
            "<?php
echo strtoupper('Abc');
echo '|';
echo strtolower('AbC');
"
        ),
        vec!["ABC|abc".to_string()]
    );
}

#[test]
fn string_title_and_case_fold_runtime() {
    assert_eq!(
        run_prints(
            "<?php
echo ucfirst('hello');
echo '|';
echo ucfirst(strtolower('world'));
"
        ),
        vec!["Hello|World".to_string()]
    );
}

#[test]
fn string_word_wrap_runtime() {
    assert_eq!(
        run_prints(
            "<?php
echo wordwrap('a b c d e', 3, '|');
"
        ),
        vec!["a b|c d|e".to_string()]
    );
}

#[test]
fn string_chunk_split_runtime() {
    assert_eq!(
        run_prints(
            "<?php
echo chunk_split('abcdef', 2, '-');
"
        ),
        vec!["ab-cd-ef-".to_string()]
    );
}

#[test]
fn string_parse_and_query_runtime() {
    assert_eq!(
        run_prints(
            "<?php
$query = 'a=1&b=2&c=three';
parse_str($query, $out);
ksort($out);
echo $out['a'] . '|' . $out['c'];
"
        ),
        vec!["1|three".to_string()]
    );
}

#[test]
fn string_explode_and_implode_runtime() {
    assert_eq!(
        run_prints(
            "<?php
$parts = explode('|', 'one|two|three');
echo implode('-', $parts);
"
        ),
        vec!["one-two-three".to_string()]
    );
}

#[test]
fn string_explode_limit_runtime() {
    assert_eq!(
        run_prints(
            "<?php
$parts = explode('|', 'a|b|c|d', 3);
echo count($parts);
echo '|';
echo $parts[2];
"
        ),
        vec!["3|c|d".to_string()]
    );
}

#[test]
fn string_implode_joins_empty_fragments_runtime() {
    assert_eq!(
        run_prints(
            "<?php
$parts = ['x', '', 'y', null];
echo implode(':', $parts);
"
        ),
        vec!["x::y:".to_string()]
    );
}

#[test]
fn string_pad_and_trim_runtime() {
    assert_eq!(
        run_prints(
            "<?php
echo str_pad('id', 6, '0', STR_PAD_LEFT);
echo '|';
echo trim('  spaced  ', ' ');
"
        ),
        vec!["0000id|spaced".to_string()]
    );
}

#[test]
fn string_replace_with_array_maps_runtime() {
    assert_eq!(
        run_prints(
            "<?php
echo strtr('abc', ['a' => 'A', 'b' => 'B']);
echo '|';
echo str_replace(['a', 'c'], ['x', 'z'], 'abcabc');
"
        ),
        vec!["ABc|xbzxbz".to_string()]
    );
}

#[test]
fn string_find_positions_runtime() {
    assert_eq!(
        run_prints(
            "<?php
echo strpos('hello world', 'world');
echo '|';
echo stripos('Hello', 'he');
"
        ),
        vec!["6|0".to_string()]
    );
}

#[test]
fn string_similarity_compare_runtime() {
    assert_eq!(
        run_prints(
            "<?php
echo strcasecmp('abc', 'ABC');
echo '|';
echo substr_compare('abcdef', 'BCD', 1, 3, true);
"
        ),
        vec!["0|0".to_string()]
    );
}

#[test]
fn string_split_even_chunks_runtime() {
    assert_eq!(
        run_prints(
            "<?php
echo json_encode(str_split('abcdef', 2));
echo '|';
echo implode('', str_split('xy'));
"
        ),
        vec!["[\"ab\",\"cd\",\"ef\"]|xy".to_string()]
    );
}

#[test]
fn string_pattern_matching_runtime() {
    assert_eq!(
        run_prints(
            "<?php
echo str_starts_with('filesystem', 'file');
echo '|';
echo str_ends_with('filesystem', 'em');
"
        ),
        vec!["1|1".to_string()]
    );
}

#[test]
fn string_html_transform_runtime() {
    assert_eq!(
        run_prints(
            "<?php
echo htmlspecialchars('<b>safe</b>');
echo '|';
echo html_entity_decode('&lt;b&gt;');
"
        ),
        vec!["&lt;b&gt;safe&lt;/b&gt;|<b>".to_string()]
    );
}

#[test]
fn string_implode_empty_values_runtime() {
    let out = run_prints(
        "<?php
echo implode('|', []);
echo '|';
echo implode('', ['x', '', 'y']);
echo '|';
echo implode('-', [1, null, 2]);
",
    );
    assert_eq!(out, vec!["|x|1--2".to_string()]);
}

#[test]
fn string_explode_no_limit_runtime() {
    let out = run_prints(
        "<?php
$parts = explode('|', 'x|y|z');
echo count($parts);
echo '|';
echo $parts[1];
",
    );
    assert_eq!(out, vec!["3|y".to_string()]);
}

#[test]
fn string_explode_limit_zero_behavior_runtime() {
    let out = run_prints(
        "<?php
$parts = explode('|', 'a|b|c', 0);
echo count($parts);
echo '|';
echo $parts[0];
",
    );
    assert_eq!(out, vec!["1|a|b|c".to_string()]);
}

#[test]
fn string_explode_limit_negative_behavior_runtime() {
    let out = run_prints(
        "<?php
$parts = explode('|', 'a|b|c', -1);
echo count($parts);
echo '|';
echo $parts[1];
",
    );
    assert_eq!(out, vec!["2|b".to_string()]);
}

#[test]
fn string_explode_with_empty_delimiter_runtime() {
    let out = run_prints(
        "<?php
$parts = str_split('abc', 1);
echo count($parts);
echo '|';
echo $parts[2];
",
    );
    assert_eq!(out, vec!["3|c".to_string()]);
}

#[test]
fn string_implode_with_nested_array_error_runtime() {
    let out = run_prints(
        "<?php
echo implode(',', [['a', 'b'], ['c']]);
",
    );
    assert_eq!(out, vec!["Array,Array".to_string()]);
}

#[test]
fn string_replace_count_param_runtime() {
    let out = run_prints(
        "<?php
$replaced = str_replace('a', 'b', 'aaxa', 2);
echo $replaced;
",
    );
    assert_eq!(out, vec!["bbxa".to_string()]);
}

#[test]
fn string_preg_split_word_runtime() {
    let out = run_prints(
        "<?php
$parts = preg_split('/\\s+/', 'one  two   three');
echo count($parts);
echo '|';
echo $parts[0];
echo '|';
echo $parts[2];
",
    );
    assert_eq!(out, vec!["3|one|three".to_string()]);
}

#[test]
fn string_preg_split_no_empty_runtime() {
    let out = run_prints(
        "<?php
$parts = preg_split('/\\s+/', 'a  b c', -1, PREG_SPLIT_NO_EMPTY);
echo count($parts);
echo '|';
echo $parts[1];
",
    );
    assert_eq!(out, vec!["3|b".to_string()]);
}

#[test]
fn string_substr_replace_range_runtime() {
    let out = run_prints(
        "<?php
echo substr_replace('abcdef', 'ZZ', 1, 3);
echo '|';
echo substr_replace('abcdef', 'QQ', -2);
",
    );
    assert_eq!(out, vec!["aZZdef|abcdQQ".to_string()]);
}

#[test]
fn string_tokenize_with_strtok_runtime() {
    let out = run_prints(
        "<?php
echo strtok('a,b,c', ',');
echo '|';
echo strtok(',');
echo '|';
echo strtok(',', 'x|y|z');
",
    );
    assert_eq!(out, vec!["a|b|x".to_string()]);
}

#[test]
fn string_explode_three_way_limit_runtime() {
    let out = run_prints(
        "<?php
$parts = explode('|', 'x|y|z|w', 3);
echo count($parts);
echo '|';
echo $parts[0];
echo '|';
echo $parts[1];
echo '|';
echo $parts[2];
",
    );
    assert_eq!(out, vec!["3|x|y|z|w".to_string()]);
}

#[test]
fn string_explode_negative_limit_keep_last_runtime() {
    let out = run_prints(
        "<?php
$parts = explode('|', 'x|y|z|', -1);
echo count($parts);
echo '|';
echo $parts[0];
echo '|';
echo $parts[1];
echo '|';
echo $parts[2];
",
    );
    assert_eq!(out, vec!["3|x|y|z".to_string()]);
}

#[test]
fn string_implode_map_values_runtime() {
    let out = run_prints(
        "<?php
$items = ['a' => 1, 'b' => null, 'c' => 'x', 'd' => false];
echo implode('-', array_values($items));
",
    );
    assert_eq!(out, vec!["1--x-".to_string()]);
}

#[test]
fn string_strip_tags_allow_list_runtime() {
    let out = run_prints(
        "<?php
echo strip_tags('<b>hi</b><i>there</i>', '<b>');
",
    );
    assert_eq!(out, vec!["<b>hi</b>there".to_string()]);
}

#[test]
fn string_addcslashes_runtime() {
    let out = run_prints(
        "<?php
echo addcslashes('a b c', ' a');
echo '|';
echo addcslashes(\"x\\n\", \"\\n\");
",
    );
    assert_eq!(out, vec!["a\\ b\\ c|x\\n".to_string()]);
}

#[test]
fn string_replace_case_insensitive_runtime() {
    let out = run_prints(
        "<?php
echo str_ireplace('WORLD', 'PHP', 'hello world');
echo '|';
echo str_ireplace(['A', 'B'], ['X', 'Y'], 'ab');
",
    );
    assert_eq!(out, vec!["hello PHP|XY".to_string()]);
}

#[test]
fn string_parse_str_nested_runtime() {
    let out = run_prints(
        "<?php
$query = 'user[name]=alice&user[id]=7&tags[]=a&tags[]=b';
parse_str($query, $out);
echo $out['user']['name'];
echo '|';
echo $out['user']['id'];
echo '|';
echo implode(',', $out['tags']);
",
    );
    assert_eq!(out, vec!["alice|7|a,b".to_string()]);
}

#[test]
fn string_nl2br_mode_runtime() {
    let out = run_prints(
        "<?php
echo nl2br(\"a\\n\", false);
echo '|';
echo nl2br(\"a\\r\\n\", true);
",
    );
    assert_eq!(out, vec!["a<br />\n|a<br>\r\n".to_string()]);
}

#[test]
fn string_offset_assign_runtime() {
    let out = run_prints(
        "<?php
$s = 'abc';
$s[1] = 'Z';
echo $s;
echo '|';
$bytes = strlen($s);
echo $bytes;
",
    );
    assert_eq!(out, vec!["aZc|3".to_string()]);
}

#[test]
fn string_interpolation_nested_index_runtime() {
    assert_eq!(
        run_prints(
            "<?php
$payload = ['user' => ['name' => 'Alice', 'roles' => ['admin', 'editor']]];
echo \"{$payload['user']['name']}=>{$payload['user']['roles'][1]}\";
"
        ),
        vec!["Alice=>editor".to_string()]
    );
}

#[test]
fn string_interpolation_variable_variable_runtime() {
    assert_eq!(
        run_prints(
            "<?php
$key = 'target';
$target = 'ok';
echo \"value:$${key}\";
"
        ),
        vec!["value:ok".to_string()]
    );
}

#[test]
fn string_heredoc_with_interpolation_runtime() {
    assert_eq!(
        run_prints(
            "<?php
$tag = 'OK';
$s = <<<TXT\n[$tag] line\nTXT;
echo trim($s);
"
        ),
        vec!["[OK] line".to_string()]
    );
}

#[test]
fn string_nowdoc_no_interpolation_runtime() {
    assert_eq!(
        run_prints(
            "<?php
$tag = 'X';
$s = <<<'TXT'
No $tag interpolation
TXT;
echo trim($s);
"
        ),
        vec!["No $tag interpolation".to_string()]
    );
}

#[test]
fn string_concat_parentheses_runtime() {
    assert_eq!(
        run_prints(
            "<?php
$left = 'a';
$right = 'c';
echo $left . ('b' . $right);
"
        ),
        vec!["abc".to_string()]
    );
}

#[test]
fn string_concat_with_ternary_runtime() {
    assert_eq!(
        run_prints(
            "<?php
echo 'x' . (true ? 'y' : 'z') . 'z';
"
        ),
        vec!["xyz".to_string()]
    );
}

#[test]
fn string_explode_empty_subject_runtime() {
    assert_eq!(
        run_prints(
            "<?php
$parts = explode('|', '');
echo count($parts);
echo '|';
echo $parts[0] === '' ? 'empty' : 'full';
"
        ),
        vec!["1|empty".to_string()]
    );
}

#[test]
fn string_explode_limit_one_runtime() {
    assert_eq!(
        run_prints(
            "<?php
$parts = explode('|', 'a|b|c', 1);
echo count($parts);
echo '|';
echo $parts[0];
"
        ),
        vec!["1|a|b|c".to_string()]
    );
}

#[test]
fn string_strip_tags_complex_input_runtime() {
    assert_eq!(
        run_prints(
            "<?php
$html = '<div><span>safe</span> &amp; <script>bad()</script></div>';
echo strip_tags($html, '<div><span>');
"
        ),
        vec!["<div><span>safe</span> &amp; ".to_string()]
    );
}

#[test]
fn string_strcmp_comparison_runtime() {
    assert_eq!(
        run_prints(
            "<?php
echo strcmp('apple', 'banana');
echo '|';
echo strcmp('apple', 'apple');
"
        ),
        vec!["-1|0".to_string()]
    );
}

#[test]
fn string_strlen_empty_and_unicode_bytes_runtime() {
    assert_eq!(
        run_prints(
            "<?php
echo strlen('');
echo '|';
echo strlen('é');
"
        ),
        vec!["0|2".to_string()]
    );
}

#[test]
fn string_strspn_and_strcspn_runtime() {
    assert_eq!(
        run_prints(
            "<?php
echo strspn('abcdef', 'abc');
echo '|';
echo strcspn('abcdef', 'de');
"
        ),
        vec!["3|3".to_string()]
    );
}

#[test]
fn string_strpos_negative_offset_runtime() {
    assert_eq!(
        run_prints(
            "<?php
echo strpos('hello world', 'o', -5);
echo '|';
echo strpos('hello', 'x', -3);
"
        ),
        vec!["7|".to_string()]
    );
}

#[test]
fn string_getcsv_quoted_fields_runtime() {
    assert_eq!(
        run_prints(
            "<?php
$row = str_getcsv('\"a\",\"b,b\",\"c\"');
echo count($row);
echo '|';
echo $row[1];
"
        ),
        vec!["3|b,b".to_string()]
    );
}

#[test]
fn string_strtr_array_overlap_runtime() {
    assert_eq!(
        run_prints(
            "<?php
echo strtr('abc', ['ab' => 'xy', 'a' => 'x']);
"
        ),
        vec!["xyc".to_string()]
    );
}

#[test]
fn string_empty_needle_compat_runtime() {
    assert_eq!(
        run_prints(
            "<?php
echo str_starts_with('abc', '');
echo '|';
echo str_ends_with('abc', '');
echo '|';
echo str_contains('', 'a');
"
        ),
        vec!["1|1|0".to_string()]
    );
}
