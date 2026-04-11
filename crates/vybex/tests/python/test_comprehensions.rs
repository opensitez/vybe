use super::helpers::*;

// ══════════════════════════════════════════════════════════════════════════════
// List comprehension runtime
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn list_comp_squares() {
    compile_ok("result = [x * x for x in range(5)]\n");
}

#[test]
fn list_comp_filtered() {
    compile_ok("result = [x for x in range(10) if x % 2 == 0]\n");
}

#[test]
fn list_comp_multiple_conditions() {
    compile_ok("result = [x for x in range(100) if x % 2 == 0 if x % 3 == 0]\n");
}

#[test]
fn list_comp_nested() {
    compile_ok("flat = [x for row in [[1,2],[3,4]] for x in row]\n");
}

#[test]
fn list_comp_matrix() {
    compile_ok("matrix = [[i*j for j in range(3)] for i in range(3)]\n");
}

#[test]
fn list_comp_with_call() {
    compile_ok("upper = [s.upper() for s in ['a', 'b', 'c']]\n");
}

// ══════════════════════════════════════════════════════════════════════════════
// Dict comprehension
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn dict_comp_basic() {
    compile_ok("d = {k: k*2 for k in range(5)}\n");
}

#[test]
fn dict_comp_from_list() {
    compile_ok("d = {s: len(s) for s in ['hello', 'world']}\n");
}

#[test]
fn dict_comp_filtered() {
    compile_ok("d = {k: v for k, v in items.items() if v > 0}\n");
}

// ══════════════════════════════════════════════════════════════════════════════
// Set comprehension
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn set_comp_basic() {
    compile_ok("s = {x * x for x in range(5)}\n");
}

#[test]
fn set_comp_from_string() {
    compile_ok("s = {c for c in 'hello world' if c != ' '}\n");
}

// ══════════════════════════════════════════════════════════════════════════════
// Generator expressions
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn generator_in_sum() {
    compile_ok("total = sum(x * x for x in range(10))\n");
}

#[test]
fn generator_in_any() {
    compile_ok("any_big = any(x > 100 for x in data)\n");
}

#[test]
fn generator_in_join() {
    compile_ok("joined = ','.join(str(x) for x in [1,2,3])\n");
}

// ══════════════════════════════════════════════════════════════════════════════
// Comprehension with ternary
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn comp_with_ternary() {
    compile_ok("result = ['even' if x % 2 == 0 else 'odd' for x in range(5)]\n");
}

#[test]
fn comp_with_method_chain() {
    compile_ok("result = [word.strip().lower() for word in lines]\n");
}
