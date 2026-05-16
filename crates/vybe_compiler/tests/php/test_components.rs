use super::helpers::compile_ok;

// These tests originally exercised the `vybe_compiler_php::compile_component`
// API, which produced a `Component` with name/exports/etc. The vybex pipeline
// doesn't have a per-language component shim — components are built one layer
// up. Here we just verify the source compiles cleanly under the unified
// pipeline; the component-shape assertions move to vybex's component layer.

// ── Component compilation ───────────────────────────────────
#[test]
fn compile_component_basic() {
    compile_ok(r#"<?php
function greet($name) {
    return 'Hello ' . $name;
}
echo greet('World');
"#);
}

#[test]
fn compile_component_with_class() {
    compile_ok(r#"<?php
class Calculator {
    public function add($a, $b) { return $a + $b; }
    public function sub($a, $b) { return $a - $b; }
}
"#);
}

#[test]
fn compile_component_exports_functions() {
    compile_ok(r#"<?php
function add($a, $b) { return $a + $b; }
function multiply($a, $b) { return $a * $b; }
"#);
}

#[test]
fn compile_component_with_traits_and_interfaces() {
    compile_ok(r#"<?php
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
}

// ── Cross-language exception compatibility ──────────────────
#[test]
fn canonical_exception_names() {
    use vybe_compiler::emitter::errors::canonical_exception_name;
    assert_eq!(canonical_exception_name("RuntimeException"), "RuntimeError");
    assert_eq!(canonical_exception_name("ValueError"), "ValueError");
    assert_eq!(canonical_exception_name("TypeError"), "TypeError");
    assert_eq!(canonical_exception_name("IndexOutOfRangeException"), "IndexError");
    assert_eq!(canonical_exception_name("Exception"), "Exception");
    assert_eq!(canonical_exception_name("Error"), "Exception");
}

#[test]
fn exception_cross_compat() {
    compile_ok(r#"<?php
try {
    throw new RuntimeException('something broke');
} catch (RuntimeException $e) {
    echo $e;
}
"#);
}

#[test]
fn multi_catch_types() {
    compile_ok(r#"<?php
try {
    throw new Exception('oops');
} catch (TypeError | ValueError $e) {
    echo 'type or value error';
} catch (Exception $e) {
    echo 'generic';
}
"#);
}
