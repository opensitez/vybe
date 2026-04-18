use vybec::parser_php::Parser;

fn parse(src: &str) -> vybec::parser_php::Program {
    Parser::new(src).expect("lexer failed").parse_program().expect("parse failed")
}

// ── Component compilation ───────────────────────────────────
#[test]
fn compile_component_basic() {
    let program = parse(r#"<?php
function greet($name) {
    return 'Hello ' . $name;
}
echo greet('World');
"#);
    let component = vybec::compiler_php::compile_component(&program, "greeting").unwrap();
    assert_eq!(component.name, "greeting");
    assert_eq!(component.language, vybe_bytecode::component::Language::Php);
    assert!(component.chunks.len() >= 2); // script + greet function
}

#[test]
fn compile_component_with_class() {
    let program = parse(r#"<?php
class Calculator {
    public function add($a, $b) { return $a + $b; }
    public function sub($a, $b) { return $a - $b; }
}
"#);
    let component = vybec::compiler_php::compile_component(&program, "math").unwrap();
    assert_eq!(component.name, "math");
    // Should have chunks for script + class constructor + methods
    assert!(component.chunks.len() >= 3);
}

#[test]
fn compile_component_exports_functions() {
    let program = parse(r#"<?php
function add($a, $b) { return $a + $b; }
function multiply($a, $b) { return $a * $b; }
"#);
    let component = vybec::compiler_php::compile_component(&program, "math_utils").unwrap();
    // Exported functions should be discoverable
    assert!(!component.exports.is_empty());
}

#[test]
fn compile_component_with_traits_and_interfaces() {
    let program = parse(r#"<?php
interface Printable {
    public function toString(): string;
}
trait Loggable {
    public function log() { echo $this->toString(); }
}
class Item implements Printable {
    use Loggable;
    public $name;
    public function __construct($name) { $this->name = $name; }
    public function toString(): string { return $this->name; }
}
"#);
    let component = vybec::compiler_php::compile_component(&program, "items").unwrap();
    assert_eq!(component.name, "items");
    assert!(component.chunks.len() >= 2);
}

// ── Cross-language exception compatibility ──────────────────
#[test]
fn canonical_exception_names() {
    use vybe_compiler_common::errors::canonical_exception_name;
    // PHP → canonical (same as Python/JS/Dart would use)
    assert_eq!(canonical_exception_name("RuntimeException"), "RuntimeError");
    assert_eq!(canonical_exception_name("ValueError"), "ValueError");
    assert_eq!(canonical_exception_name("TypeError"), "TypeError");
    assert_eq!(canonical_exception_name("IndexOutOfRangeException"), "IndexError");
    assert_eq!(canonical_exception_name("Exception"), "Exception");
    assert_eq!(canonical_exception_name("Error"), "Exception");
}

#[test]
fn exception_cross_compat() {
    // PHP throw → Python catch should work because both use canonical names
    let program = parse(r#"<?php
try {
    throw new RuntimeException('something broke');
} catch (RuntimeException $e) {
    echo $e;
}
"#);
    let chunks = vybec::compiler_php::Compiler::new().compile(&program).unwrap();
    assert!(!chunks.is_empty());
}

#[test]
fn multi_catch_types() {
    let program = parse(r#"<?php
try {
    throw new Exception('oops');
} catch (TypeError | ValueError $e) {
    echo 'type or value error';
} catch (Exception $e) {
    echo 'generic';
}
"#);
    let chunks = vybec::compiler_php::Compiler::new().compile(&program).unwrap();
    assert!(!chunks.is_empty());
}
