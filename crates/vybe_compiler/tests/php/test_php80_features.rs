use super::helpers::run_prints;

// ── Match expression (PHP 8.0) ────────────────────────────────

#[test] fn match_strict_type_check() {
    assert_eq!(run_prints(r#"<?php
echo match(0) { false => 'false', null => 'null', 0 => 'zero', default => 'other' };
"#), vec!["zero"]);
}
#[test] fn match_as_statement() {
    assert_eq!(run_prints(r#"<?php
$x = 3;
match($x) { 1 => print('one'), 2 => print('two'), 3 => print('three') };
"#), vec!["three"]);
}

// ── Nullsafe operator (PHP 8.0) ───────────────────────────────

#[test] fn nullsafe_chain_with_method() {
    assert_eq!(run_prints(r#"<?php
class Session { private ?User2 $user = null; public function getUser(): ?User2 { return $this->user; } }
class User2 { public function getAddress(): ?Addr { return null; } }
class Addr { public function getCity(): string { return 'Paris'; } }
$session = new Session;
echo $session->getUser()?->getAddress()?->getCity() ?? 'no city';
"#), vec!["no city"]);
}

// ── Named arguments (PHP 8.0) ─────────────────────────────────

#[test] fn named_arg_basic() {
    assert_eq!(run_prints(r#"<?php
function createTag(string $tag, string $content, string $class = ''): string {
    $cls = $class ? " class=\"$class\"" : '';
    return "<$tag$cls>$content</$tag>";
}
echo createTag(content: 'Hello', tag: 'p', class: 'greeting');
"#), vec!["<p class=\"greeting\">Hello</p>"]);
}

// ── Union types (PHP 8.0) ─────────────────────────────────────

#[test] fn union_type_int_string_param() {
    assert_eq!(run_prints(r#"<?php
function padId(int|string $id, int $len = 5): string { return str_pad((string)$id, $len, '0', STR_PAD_LEFT); }
echo padId(42) . ',' . padId('7');
"#), vec!["00042,00007"]);
}

// ── Attributes (PHP 8.0) ─────────────────────────────────────

#[test] fn attribute_on_class() {
    assert_eq!(run_prints(r#"<?php
#[Attribute]
class Route { public function __construct(public string $path) {} }
#[Route('/home')]
class HomeController {}
$ref = new ReflectionClass(HomeController::class);
$attrs = $ref->getAttributes(Route::class);
echo $attrs[0]->newInstance()->path;
"#), vec!["/home"]);
}
#[test] fn attribute_on_method() {
    assert_eq!(run_prints(r#"<?php
#[Attribute]
class Deprecated { public function __construct(public string $since) {} }
class OldApi {
    #[Deprecated(since: '2.0')]
    public function oldMethod(): void {}
}
$ref = new ReflectionMethod(OldApi::class, 'oldMethod');
$attrs = $ref->getAttributes(Deprecated::class);
echo $attrs[0]->newInstance()->since;
"#), vec!["2.0"]);
}
#[test] fn attribute_on_property() {
    assert_eq!(run_prints(r#"<?php
#[Attribute]
class Column { public function __construct(public string $name) {} }
class User3 {
    #[Column('user_name')]
    public string $username = '';
}
$ref = new ReflectionProperty(User3::class, 'username');
echo $ref->getAttributes(Column::class)[0]->newInstance()->name;
"#), vec!["user_name"]);
}

// ── throw expression (PHP 8.0) ────────────────────────────────

#[test] fn throw_in_ternary() {
    assert_eq!(run_prints(r#"<?php
function nonEmpty(string $s): string {
    return strlen($s) > 0 ? $s : throw new InvalidArgumentException('empty');
}
try { echo nonEmpty('') ; } catch (InvalidArgumentException $e) { echo 'caught'; }
echo ',' . nonEmpty('hi');
"#), vec!["caught,hi"]);
}
#[test] fn throw_in_null_coalesce() {
    assert_eq!(run_prints(r#"<?php
function getOrThrow(?string $val): string {
    return $val ?? throw new \RuntimeException('null');
}
try { getOrThrow(null); } catch (\RuntimeException $e) { echo $e->getMessage(); }
"#), vec!["null"]);
}
#[test] fn throw_in_arrow_fn() {
    assert_eq!(run_prints(r#"<?php
$validate = fn($n) => $n > 0 ? $n : throw new RangeException("Expected positive, got $n");
try { $validate(-1); } catch (RangeException $e) { echo $e->getMessage(); }
"#), vec!["Expected positive, got -1"]);
}

// ── str_contains / str_starts_with / str_ends_with (PHP 8.0) ──

#[test] fn str_contains_false() {
    assert_eq!(run_prints(r#"<?php echo str_contains('Hello World', 'xyz') ? 'yes' : 'no'; "#), vec!["no"]);
}
#[test] fn str_starts_with_false() {
    assert_eq!(run_prints(r#"<?php echo str_starts_with('Hello World', 'World') ? 'yes' : 'no'; "#), vec!["no"]);
}
#[test] fn str_ends_with_false() {
    assert_eq!(run_prints(r#"<?php echo str_ends_with('Hello World', 'Hello') ? 'yes' : 'no'; "#), vec!["no"]);
}

// ── fdiv (PHP 8.0) ────────────────────────────────────────────

#[test] fn fdiv_inf_result() {
    assert_eq!(run_prints(r#"<?php echo fdiv(10, 0); "#), vec!["INF"]);
}
#[test] fn fdiv_nan_result() {
    assert_eq!(run_prints(r#"<?php echo fdiv(0, 0); "#), vec!["NAN"]);
}

// ── get_debug_type (PHP 8.0) ──────────────────────────────────

#[test] fn get_debug_type_int() {
    assert_eq!(run_prints(r#"<?php echo get_debug_type(42); "#), vec!["int"]);
}
#[test] fn get_debug_type_object() {
    assert_eq!(run_prints(r#"<?php class Foo {} echo get_debug_type(new Foo); "#), vec!["Foo"]);
}
#[test] fn get_debug_type_array() {
    assert_eq!(run_prints(r#"<?php echo get_debug_type([]); "#), vec!["array"]);
}
