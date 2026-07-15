use super::helpers::{compile_ok, run_prints};

fn assert_outputs(src: &str, expected: &[&str]) {
    assert_eq!(run_prints(src), expected);
}

#[test]
fn forward_function_call_inside_function_runtime() {
    assert_outputs(
        "<?php print f(); function f(){ return g(); } function g(){ return 'G'; }",
        &["G"],
    );
}

#[test]
fn function_call_survives_same_named_variable_runtime() {
    assert_outputs(
        "<?php function translate($s){ return $s; } $translate = ['x' => 'y']; echo translate('ok');",
        &["ok"],
    );
}

#[test]
fn case_insensitive_in_array_keeps_php_argument_order() {
    assert_outputs("<?php echo in_Array('en', ['en']) ? 'yes' : 'no';", &["yes"]);
}

// ── String builtins ─────────────────────────────────────────
#[test]
fn strlen() {
    assert_outputs("<?php echo strlen('hello');", &["5"]);
}
#[test]
fn strtolower() {
    assert_outputs("<?php echo strtolower('HELLO');", &["hello"]);
}
#[test]
fn strtoupper() {
    assert_outputs("<?php echo strtoupper('hello');", &["HELLO"]);
}
#[test]
fn trim() {
    assert_outputs("<?php echo trim('  hello  ');", &["hello"]);
}
#[test]
fn ltrim() {
    assert_outputs("<?php echo ltrim('  hello');", &["hello"]);
}
#[test]
fn rtrim() {
    assert_outputs("<?php echo rtrim('hello  ');", &["hello"]);
}
#[test]
fn substr() {
    assert_outputs("<?php echo substr('hello', 1, 3);", &["ell"]);
}
#[test]
fn str_replace() {
    assert_outputs("<?php echo str_replace('o', '0', 'hello');", &["hell0"]);
}
#[test]
fn explode() {
    assert_outputs(
        "<?php echo json_encode(explode(',', 'a,b,c'));",
        &["[\"a\",\"b\",\"c\"]"],
    );
}
#[test]
fn implode() {
    assert_outputs("<?php echo implode(',', ['a','b','c']);", &["a,b,c"]);
}
#[test]
fn strpos() {
    assert_outputs("<?php echo strpos('hello', 'lo');", &["3"]);
}
#[test]
fn str_contains() {
    assert_outputs(
        "<?php echo str_contains('hello', 'ell') ? 'yes' : 'no';",
        &["yes"],
    );
}
#[test]
fn str_starts_with() {
    assert_outputs(
        "<?php echo str_starts_with('hello', 'he') ? 'yes' : 'no';",
        &["yes"],
    );
}
#[test]
fn str_ends_with() {
    assert_outputs(
        "<?php echo str_ends_with('hello', 'lo') ? 'yes' : 'no';",
        &["yes"],
    );
}
#[test]
fn superglobal_string_key_read_after_assignment() {
    assert_outputs(
        "<?php $_SERVER = ['SCRIPT_NAME' => '/cgi-bin/php.cgi']; echo $_SERVER['SCRIPT_NAME'];",
        &["/cgi-bin/php.cgi"],
    );
}
#[test]
fn str_contains_dynamic_array_value() {
    assert_outputs(
        "<?php $_SERVER = ['SCRIPT_NAME' => '/cgi-bin/php.cgi']; echo str_contains($_SERVER['SCRIPT_NAME'], 'php.cgi') ? 'yes' : 'no';",
        &["yes"],
    );
}
#[test]
fn str_ends_with_dynamic_array_value() {
    assert_outputs(
        "<?php $_SERVER = ['SCRIPT_FILENAME' => '/usr/bin/php.cgi']; echo str_ends_with($_SERVER['SCRIPT_FILENAME'], 'php.cgi') ? 'yes' : 'no';",
        &["yes"],
    );
}
#[test]
fn str_repeat() {
    assert_outputs("<?php echo str_repeat('ab', 3);", &["ababab"]);
}
#[test]
fn str_pad() {
    assert_outputs("<?php echo str_pad('42', 5, '0');", &["42000"]);
}
#[test]
fn chr_ord() {
    assert_outputs(
        "<?php echo chr(65), \"\\n\"; echo ord('A'), \"\\n\";",
        &["A", "65"],
    );
}
#[test]
fn ucfirst() {
    assert_outputs("<?php echo ucfirst('hello');", &["Hello"]);
}
#[test]
fn lcfirst() {
    assert_outputs("<?php echo lcfirst('Hello');", &["hello"]);
}
#[test]
fn string_post_inc_trailing_digits() {
    assert_outputs("<?php $x = '2026-03-25'; $x++; echo $x;", &["2026-03-26"]);
}
#[test]
fn string_post_inc_date_loop_terminates() {
    assert_outputs(
        "<?php $x = '2026-03-25'; $e = '2026-03-31'; $count = 0; while ($x <= $e && $count < 20) { $count = $count + 1; $x++; } echo $count;",
        &["7"],
    );
}
#[test]
fn nl2br() {
    compile_ok("<?php $x = nl2br(\"hello\\nworld\");");
}
#[test]
fn htmlspecialchars() {
    assert_outputs(
        "<?php echo htmlspecialchars('<b>hi</b>');",
        &["&lt;b&gt;hi&lt;/b&gt;"],
    );
}
#[test]
fn sprintf() {
    compile_ok("<?php $x = sprintf('Hello %s, age %d', 'John', 30);");
}

// ── Array builtins ──────────────────────────────────────────
#[test]
fn count() {
    assert_outputs("<?php echo count([1,2,3]);", &["3"]);
}
#[test]
fn array_push() {
    assert_outputs(
        "<?php $a = [1]; array_push($a, 2); echo json_encode($a);",
        &["[1,2]"],
    );
}
#[test]
fn array_pop() {
    assert_outputs(
        "<?php $a = [1,2]; $v = array_pop($a); echo $v, \"\\n\"; echo json_encode($a), \"\\n\";",
        &["2", "[1]"],
    );
}
#[test]
fn array_shift() {
    assert_outputs(
        "<?php $a = [1,2]; $v = array_shift($a); echo $v, \"\\n\"; echo json_encode($a), \"\\n\";",
        &["1", "[2]"],
    );
}
#[test]
fn array_reverse() {
    assert_outputs(
        "<?php echo json_encode(array_reverse([1,2,3]));",
        &["[3,2,1]"],
    );
}
#[test]
fn array_slice() {
    assert_outputs(
        "<?php echo json_encode(array_slice([1,2,3,4], 1, 2));",
        &["[2,3]"],
    );
}
#[test]
fn array_merge() {
    assert_outputs(
        "<?php echo json_encode(array_merge([1,2], [3,4]));",
        &["[1,2,3,4]"],
    );
}
#[test]
fn array_search() {
    assert_outputs("<?php echo array_search(2, [1,2,3]);", &["1"]);
}
#[test]
fn in_array() {
    assert_outputs("<?php echo in_array(2, [1,2,3]) ? 'yes' : 'no';", &["yes"]);
}
#[test]
fn array_keys() {
    compile_ok("<?php $x = array_keys(['a'=>1,'b'=>2]);");
}
#[test]
fn array_values() {
    compile_ok("<?php $x = array_values(['a'=>1,'b'=>2]);");
}
#[test]
fn sort() {
    assert_outputs(
        "<?php $a = [3,1,2]; sort($a); echo json_encode($a);",
        &["[1,2,3]"],
    );
}
#[test]
fn range() {
    assert_outputs("<?php echo json_encode(range(1, 4));", &["[1,2,3,4]"]);
}
#[test]
fn array_sum() {
    assert_outputs("<?php echo array_sum([1,2,3]);", &["6"]);
}
#[test]
fn compact() {
    compile_ok("<?php $a = 1; $b = 2; $x = compact('a', 'b');");
}
#[test]
fn array_key_exists() {
    assert_outputs(
        "<?php echo array_key_exists('a', ['a'=>1]) ? 'yes' : 'no';",
        &["yes"],
    );
}

// ── Callback array ops ──────────────────────────────────────
#[test]
fn array_map() {
    assert_outputs(
        "<?php echo json_encode(array_map(fn($n) => $n * 2, [1,2,3]));",
        &["[2,4,6]"],
    );
}
#[test]
fn array_filter_no_cb() {
    assert_outputs(
        "<?php $x = array_filter([1,0,2,null,3]); echo count($x);",
        &["3"],
    );
}
#[test]
fn array_filter_cb() {
    assert_outputs(
        "<?php $x = array_filter([1,2,3,4], fn($n) => $n > 2); echo count($x);",
        &["2"],
    );
}
#[test]
fn array_reduce() {
    assert_outputs(
        "<?php echo array_reduce([1,2,3], fn($c,$i) => $c + $i, 0);",
        &["6"],
    );
}
#[test]
fn array_walk() {
    compile_ok("<?php $a = [1,2,3]; array_walk($a, fn($v,$k) => $v);");
}
#[test]
fn usort() {
    compile_ok("<?php $a = [3,1,2]; usort($a);");
}

// ── Math builtins ───────────────────────────────────────────
#[test]
fn abs() {
    compile_ok("<?php $x = abs(-5);");
}
#[test]
fn ceil() {
    compile_ok("<?php $x = ceil(1.2);");
}
#[test]
fn floor() {
    compile_ok("<?php $x = floor(1.8);");
}
#[test]
fn round() {
    compile_ok("<?php $x = round(1.5);");
}
#[test]
fn sqrt() {
    compile_ok("<?php $x = sqrt(16);");
}
#[test]
fn pow() {
    compile_ok("<?php $x = pow(2, 8);");
}
#[test]
fn max_min() {
    compile_ok("<?php $a = max(1,2,3); $b = min(1,2,3);");
}
#[test]
fn sin_cos_tan() {
    compile_ok("<?php $a = sin(1.0); $b = cos(1.0); $c = tan(1.0);");
}
#[test]
fn log_exp() {
    compile_ok("<?php $a = log(10); $b = exp(1);");
}
#[test]
fn rand() {
    compile_ok("<?php $x = rand();");
}

// ── Type builtins ───────────────────────────────────────────
#[test]
fn intval() {
    assert_outputs("<?php echo intval('42');", &["42"]);
}
#[test]
fn floatval() {
    assert_outputs("<?php echo floatval('3.14');", &["3.14"]);
}
#[test]
fn strval() {
    assert_outputs("<?php echo strval(42);", &["42"]);
}
#[test]
fn boolval() {
    assert_outputs("<?php echo boolval(1) ? 'yes' : 'no';", &["yes"]);
}
#[test]
fn is_null() {
    assert_outputs("<?php echo is_null(null) ? 'yes' : 'no';", &["yes"]);
}
#[test]
fn is_numeric() {
    assert_outputs("<?php echo is_numeric('42') ? 'yes' : 'no';", &["yes"]);
}
#[test]
fn is_array() {
    assert_outputs("<?php echo is_array([]) ? 'yes' : 'no';", &["yes"]);
}
#[test]
fn is_string() {
    assert_outputs("<?php echo is_string('hi') ? 'yes' : 'no';", &["yes"]);
}
#[test]
fn is_int() {
    assert_outputs("<?php echo is_int(42) ? 'yes' : 'no';", &["yes"]);
}
#[test]
fn is_bool() {
    assert_outputs("<?php echo is_bool(true) ? 'yes' : 'no';", &["yes"]);
}
#[test]
fn isset() {
    assert_outputs("<?php echo isset($a) ? 'yes' : 'no';", &["no"]);
}
#[test]
fn empty() {
    compile_ok("<?php $x = empty($a);");
}
#[test]
fn gettype() {
    assert_outputs("<?php echo gettype(42);", &["integer"]);
}
#[test]
fn define_defined() {
    compile_ok("<?php define('FOO', 42); $x = defined('FOO');");
}
#[test]
fn function_exists_builtin_runtime() {
    assert_outputs(
        "<?php echo function_exists('error_reporting') ? 'yes' : 'no';",
        &["yes"],
    );
}
#[test]
fn function_exists_detects_mysqli_connect_builtin() {
    assert_outputs(
        "<?php echo function_exists('mysqli_connect') ? 'yes' : 'no';",
        &["yes"],
    );
}
#[test]
fn function_exists_missing_runtime() {
    assert_outputs(
        "<?php echo function_exists('definitely_missing_function') ? 'yes' : 'no';",
        &["no"],
    );
}
#[test]
fn function_exists_user_defined_runtime() {
    assert_outputs(
        "<?php function LocalFn() {} echo function_exists('LocalFn') ? 'yes' : 'no';",
        &["yes"],
    );
}
#[test]
fn variable_function_name_user_defined_runtime() {
    assert_outputs(
        "<?php function LocalFn($name) { echo 'hi ' . $name; } $fn = 'LocalFn'; $fn('wp');",
        &["hi wp"],
    );
}
#[test]
fn extension_loaded_runtime_condition_shape() {
    assert_outputs(
        "<?php $is_ipv6 = false; echo ($is_ipv6 && extension_loaded('mysqlnd')) ? 'yes' : 'no';",
        &["no"],
    );
}
#[test]
fn extension_loaded_mysql_surface_runtime() {
    assert_outputs(
        "<?php echo extension_loaded('mysqlnd') ? 'yes' : 'no', \"\\n\"; echo extension_loaded('mysqli') ? 'yes' : 'no', \"\\n\"; echo extension_loaded('pdo_mysql') ? 'yes' : 'no', \"\\n\"; echo extension_loaded('definitely_missing_ext') ? 'yes' : 'no', \"\\n\";",
        &["yes", "yes", "yes", "no"],
    );
}
#[test]
fn php_version_constant_and_function_runtime() {
    assert_outputs(
        "<?php echo PHP_VERSION, \"\\n\"; echo phpversion(), \"\\n\";",
        &["8.0.0", "8.0.0"],
    );
}
#[test]
fn class_exists_missing_runtime() {
    assert_outputs(
        "<?php echo class_exists('MO', false) ? 'yes' : 'no';",
        &["no"],
    );
}
#[test]
fn class_exists_declared_runtime() {
    assert_outputs(
        "<?php class MO {} echo class_exists('MO', false) ? 'yes' : 'no';",
        &["yes"],
    );
}
#[test]
fn version_compare_returns_ordering() {
    assert_outputs(
        "<?php echo version_compare('7.0.0', '8.2.0'), \"\\n\"; echo version_compare('8.2.0', '7.0.0'), \"\\n\"; echo version_compare('8.2.0', '8.2.0'), \"\\n\";",
        &["-1", "1", "0"],
    );
}
#[test]
fn version_compare_operator_runtime() {
    assert_outputs(
        "<?php echo version_compare('8.2.0', '7.0.0', '>') ? 'yes' : 'no', \"\\n\"; echo version_compare('7.0.0', '8.2.0', '>') ? 'yes' : 'no', \"\\n\"; echo version_compare('8.2.0', '8.2.0', '>=') ? 'yes' : 'no', \"\\n\";",
        &["yes", "no", "yes"],
    );
}
#[test]
fn wordpress_php_version_error_branch_runtime() {
    assert_outputs(
        "<?php $required_php_version = '8.2.0'; $wp_version = '6.5.5'; $php_version = '7.0.0'; if ( version_compare( $required_php_version, $php_version, '>' ) ) { printf('Your server is running PHP version %1$s but WordPress %2$s requires at least %3$s.', $php_version, $wp_version, $required_php_version); exit( 1 ); } echo 'after';",
        &["Your server is running PHP version 7.0.0 but WordPress 6.5.5 requires at least 8.2.0."],
    );
}
#[test]
fn class_exists() {
    compile_ok("<?php $x = class_exists('stdClass');");
}

// ── Encoding / JSON / Crypto ────────────────────────────────
#[test]
fn json_encode() {
    assert_outputs("<?php echo json_encode(['a'=>1]);", &["{\"a\":1}"]);
}
#[test]
fn json_decode() {
    assert_outputs(
        "<?php $x = json_decode('{\"a\":1}'); echo json_encode($x);",
        &["{\"a\":1}"],
    );
}
#[test]
fn urlencode() {
    compile_ok("<?php $x = urlencode('hello world');");
}
#[test]
fn urldecode() {
    compile_ok("<?php $x = urldecode('hello%20world');");
}
#[test]
fn base64_encode() {
    assert_outputs("<?php echo base64_encode('hello');", &["aGVsbG8="]);
}
#[test]
fn base64_decode() {
    assert_outputs("<?php echo base64_decode('aGVsbG8=');", &["hello"]);
}
#[test]
fn md5() {
    compile_ok("<?php $x = md5('hello');");
}
#[test]
fn sha1() {
    compile_ok("<?php $x = sha1('hello');");
}

// ── Regex ───────────────────────────────────────────────────
#[test]
fn preg_match() {
    assert_outputs("<?php echo preg_match('/\\d+/', 'abc123');", &["1"]);
}
#[test]
fn preg_replace() {
    assert_outputs("<?php echo preg_replace('/\\d/', 'X', 'a1b2');", &["aXbX"]);
}
#[test]
fn preg_split() {
    assert_outputs(
        "<?php echo json_encode(preg_split('/,/', 'a,b,c'));",
        &["[\"a\",\"b\",\"c\"]"],
    );
}

// ── Filesystem / IO ─────────────────────────────────────────
#[test]
fn file_exists() {
    compile_ok("<?php $x = file_exists('/tmp/test');");
}
#[test]
fn dirname_basename() {
    compile_ok("<?php $x = dirname('/tmp/test.txt'); $y = basename('/tmp/test.txt');");
}
#[test]
fn time() {
    compile_ok("<?php $t = time();");
}
#[test]
fn die() {
    assert_outputs("<?php die('goodbye'); echo 'after';", &["goodbye"]);
}
#[test]
fn exit_call() {
    assert_outputs("<?php echo 'before'; exit(0); echo 'after';", &["before"]);
}
