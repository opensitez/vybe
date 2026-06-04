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
