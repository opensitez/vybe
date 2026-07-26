use super::helpers::run_prints;

// ── Heredoc basic ─────────────────────────────────────────────

#[test]
fn heredoc_basic_output() {
    assert_eq!(
        run_prints(
            r#"<?php
$text = <<<EOT
Hello World
EOT;
echo $text;
"#
        ),
        vec!["Hello World"]
    );
}

#[test]
fn heredoc_variable_interpolation() {
    assert_eq!(
        run_prints(
            r#"<?php
$name = "Alice";
$msg = <<<EOT
Hello $name
EOT;
echo $msg;
"#
        ),
        vec!["Hello Alice"]
    );
}

#[test]
fn heredoc_multiline_preserves_newlines() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = <<<EOT
line one
line two
line three
EOT;
echo substr_count($s, "\n");
"#
        ),
        vec!["2"]
    );
}

#[test]
fn heredoc_expression_interpolation() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = 3;
$b = 4;
$s = <<<EOT
Result: {$a}
EOT;
echo $s;
"#
        ),
        vec!["Result: 3"]
    );
}

#[test]
fn heredoc_array_access_interpolation() {
    assert_eq!(
        run_prints(
            r#"<?php
$data = ['key' => 'value'];
$s = <<<EOT
Got: {$data['key']}
EOT;
echo $s;
"#
        ),
        vec!["Got: value"]
    );
}

#[test]
fn heredoc_object_property_interpolation() {
    assert_eq!(
        run_prints(
            r#"<?php
$obj = new stdClass();
$obj->name = "Bob";
$s = <<<EOT
Name: {$obj->name}
EOT;
echo $s;
"#
        ),
        vec!["Name: Bob"]
    );
}

#[test]
fn heredoc_as_function_argument() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strlen(<<<EOT
hello
EOT);
"#
        ),
        vec!["5"]
    );
}

// ── Nowdoc basic ─────────────────────────────────────────────

#[test]
fn nowdoc_no_interpolation() {
    assert_eq!(
        run_prints(
            r#"<?php
$name = "Alice";
$s = <<<'EOT'
Hello $name
EOT;
echo $s;
"#
        ),
        vec!["Hello $name"]
    );
}

#[test]
fn nowdoc_preserves_backslash_sequences() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = <<<'EOT'
no \n escape here
EOT;
echo $s;
"#
        ),
        vec!["no \\n escape here"]
    );
}

#[test]
fn nowdoc_curly_braces_literal() {
    assert_eq!(
        run_prints(
            r#"<?php
$x = 42;
$s = <<<'EOT'
value: {$x}
EOT;
echo $s;
"#
        ),
        vec!["value: {$x}"]
    );
}

#[test]
fn nowdoc_as_class_constant() {
    assert_eq!(
        run_prints(
            r#"<?php
class Config {
    const TEMPLATE = <<<'EOT'
raw template
EOT;
}
echo Config::TEMPLATE;
"#
        ),
        vec!["raw template"]
    );
}

#[test]
fn nowdoc_multiline_content() {
    assert_eq!(
        run_prints(
            r#"<?php
$sql = <<<'SQL'
SELECT *
FROM users
WHERE id = :id
SQL;
echo substr_count($sql, "\n");
"#
        ),
        vec!["2"]
    );
}

// ── Flexible heredoc (PHP 7.3+ indented closing marker) ───────

#[test]
fn flexible_heredoc_indented_closing() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = <<<EOT
    Hello
    World
    EOT;
echo trim($s);
"#
        ),
        vec!["Hello", "World"]
    );
}

#[test]
fn flexible_nowdoc_indented_closing() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = <<<'EOT'
    raw $text
    EOT;
echo trim($s);
"#
        ),
        vec![r"raw $text"]
    );
}

// ── Heredoc in data structures ────────────────────────────────

#[test]
fn heredoc_as_array_value() {
    assert_eq!(
        run_prints(
            r#"<?php
$arr = [
    'msg' => <<<EOT
hello
EOT,
];
echo $arr['msg'];
"#
        ),
        vec!["hello"]
    );
}

#[test]
fn heredoc_concatenation() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = <<<EOT
foo
EOT;
$b = <<<EOT
bar
EOT;
echo trim($a) . trim($b);
"#
        ),
        vec!["foobar"]
    );
}

#[test]
fn heredoc_in_return_statement() {
    assert_eq!(
        run_prints(
            r#"<?php
function greeting(string $name): string {
    return <<<EOT
Hello, $name!
EOT;
}
echo trim(greeting("World"));
"#
        ),
        vec!["Hello, World!"]
    );
}

#[test]
fn heredoc_numeric_variable_interpolation() {
    assert_eq!(
        run_prints(
            r#"<?php
$count = 42;
$s = <<<EOT
Items: $count
EOT;
echo trim($s);
"#
        ),
        vec!["Items: 42"]
    );
}

#[test]
fn heredoc_nested_array_interpolation() {
    assert_eq!(
        run_prints(
            r#"<?php
$users = [['name' => 'Alice'], ['name' => 'Bob']];
$s = <<<EOT
First: {$users[0]['name']}
EOT;
echo trim($s);
"#
        ),
        vec!["First: Alice"]
    );
}

// ── Edge cases ────────────────────────────────────────────────

#[test]
fn heredoc_empty_body() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = <<<EOT
EOT;
echo strlen($s);
"#
        ),
        vec!["0"]
    );
}

#[test]
fn heredoc_special_chars_preserved() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = <<<EOT
<tag>value & more</tag>
EOT;
echo trim($s);
"#
        ),
        vec!["<tag>value & more</tag>"]
    );
}

#[test]
fn heredoc_unicode_content() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = <<<EOT
café
EOT;
echo trim($s);
"#
        ),
        vec!["café"]
    );
}

#[test]
fn nowdoc_dollar_sign_not_interpolated() {
    assert_eq!(
        run_prints(
            r#"<?php
$price = 9.99;
$s = <<<'EOT'
Price: $price USD
EOT;
echo trim($s);
"#
        ),
        vec!["Price: $price USD"]
    );
}

#[test]
fn heredoc_in_match_expression() {
    assert_eq!(
        run_prints(
            r#"<?php
$x = 1;
$result = match($x) {
    1 => <<<EOT
one
EOT,
    default => 'other',
};
echo trim($result);
"#
        ),
        vec!["one"]
    );
}

#[test]
fn heredoc_with_tab_indentation() {
    assert_eq!(
        run_prints("<?php\n$s = <<<EOT\n\tindented\nEOT;\necho trim($s);\n"),
        vec!["indented"]
    );
}

#[test]
fn heredoc_with_interpolated_function_call() {
    assert_eq!(
        run_prints(
            r#"<?php
$name = 'alice';
$s = <<<EOT
HELLO {$name}
EOT;
echo trim($s);
"#
        ),
        vec!["HELLO alice"]
    );
}

#[test]
fn heredoc_with_object_method_in_interpolation() {
    assert_eq!(
        run_prints(
            r#"<?php
class User {
    public function __construct(private string $name) {}
    public function label(): string { return $this->name; }
}
$u = new User('Bob');
$s = <<<EOT
User: {$u->label()}
EOT;
echo trim($s);
"#
        ),
        vec!["User: Bob"]
    );
}

#[test]
fn nowdoc_multiple_lines_and_quotes() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = <<<'EOT'
'single' "double" \$value
line two
EOT;
echo str_replace("\n", "|", trim($s));
"#
        ),
        vec!["'single' \"double\" \\$value|line two"]
    );
}

#[test]
fn heredoc_preserves_trailing_spaces() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = <<<EOT
line-one  
line-two   
EOT;
echo str_replace("\n", "|", $s);
"#
        ),
        vec!["line-one  |line-two   "]
    );
}

#[test]
fn heredoc_in_array_map_context() {
    assert_eq!(
        run_prints(
            r#"<?php
$rows = [1,2];
$labels = array_map(fn($n) => <<<EOT
item-$n
EOT, $rows);
echo implode(',', $labels);
"#
        ),
        vec!["item-1,item-2"]
    );
}
