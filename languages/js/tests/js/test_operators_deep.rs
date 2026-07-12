/// Operators — bitwise ops, exponentiation, in/instanceof advanced,
/// comma operator in loops, ternary chains, typeof checks, void,
/// logical operators with non-boolean values.
use super::helpers::run_js;

// ── bitwise operators ─────────────────────────────────────────────────────────

#[test]
fn bitwise_and_or_xor() {
    assert_eq!(
        run_js(
            r#"
console.log(0b1010 & 0b1100);  // AND: 0b1000 = 8
console.log(0b1010 | 0b1100);  // OR:  0b1110 = 14
console.log(0b1010 ^ 0b1100);  // XOR: 0b0110 = 6
"#
        ),
        vec!["8", "14", "6"]
    );
}

#[test]
fn bitwise_not() {
    assert_eq!(
        run_js(
            r#"
console.log(~0);   // -1
console.log(~1);   // -2
console.log(~-1);  // 0
"#
        ),
        vec!["-1", "-2", "0"]
    );
}

#[test]
fn left_shift_right_shift() {
    assert_eq!(
        run_js(
            r#"
console.log(1 << 4);   // 16
console.log(16 >> 2);  // 4
console.log(-1 >> 1);  // -1 (sign preserved)
console.log(-1 >>> 1); // 2147483647 (unsigned)
"#
        ),
        vec!["16", "4", "-1", "2147483647"]
    );
}

#[test]
fn unsigned_right_shift() {
    assert_eq!(
        run_js(
            r#"
console.log(-1 >>> 0);  // 4294967295 (convert to uint32)
console.log(0xFFFFFFFF >>> 0); // 4294967295
"#
        ),
        vec!["4294967295", "4294967295"]
    );
}

// ── bitwise use cases ─────────────────────────────────────────────────────────

#[test]
fn bitwise_flag_manipulation() {
    assert_eq!(
        run_js(
            r#"
const READ   = 0b001;
const WRITE  = 0b010;
const EXEC   = 0b100;

let perms = READ | WRITE; // 0b011
console.log((perms & READ) !== 0);  // has READ
console.log((perms & EXEC) !== 0);  // has EXEC (no)
perms |= EXEC;                       // add EXEC
console.log((perms & EXEC) !== 0);  // has EXEC (yes)
perms &= ~WRITE;                     // remove WRITE
console.log((perms & WRITE) !== 0); // has WRITE (no)
"#
        ),
        vec!["true", "false", "true", "false"]
    );
}

#[test]
fn double_tilde_converts_to_int32() {
    assert_eq!(
        run_js(
            r#"
console.log(~~3.9);    // 3
console.log(~~-3.9);   // -3
console.log(~~"42");   // 42
"#
        ),
        vec!["3", "-3", "42"]
    );
}

// ── exponentiation ────────────────────────────────────────────────────────────

#[test]
fn exponentiation_operator() {
    assert_eq!(
        run_js(
            r#"
console.log(2 ** 10);
console.log(3 ** 3);
console.log((-2) ** 3);
"#
        ),
        vec!["1024", "27", "-8"]
    );
}

#[test]
fn exponentiation_assignment() {
    assert_eq!(
        run_js(
            r#"
let x = 2;
x **= 8;
console.log(x);
"#
        ),
        vec!["256"]
    );
}

// ── in operator ───────────────────────────────────────────────────────────────

#[test]
fn in_with_array_index() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, 3];
console.log(0 in arr);
console.log(2 in arr);
console.log(3 in arr);  // out of bounds
"#
        ),
        vec!["true", "true", "false"]
    );
}

#[test]
fn in_with_string_key() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: undefined }; // key exists but value undefined
console.log("a" in obj);
console.log("b" in obj);
console.log(obj.a === undefined); // exists but undefined
"#
        ),
        vec!["true", "false", "true"]
    );
}

// ── instanceof advanced ───────────────────────────────────────────────────────

#[test]
fn instanceof_custom_symbol_hasinstance() {
    assert_eq!(
        run_js(
            r#"
class EvenChecker {
    static [Symbol.hasInstance](n) {
        return typeof n === "number" && n % 2 === 0;
    }
}
console.log(2 instanceof EvenChecker);
console.log(3 instanceof EvenChecker);
console.log(4 instanceof EvenChecker);
"#
        ),
        vec!["true", "false", "true"]
    );
}

// ── logical operators with non-booleans ───────────────────────────────────────

#[test]
fn logical_and_returns_first_falsy_or_last() {
    assert_eq!(
        run_js(
            r#"
console.log(1 && 2 && 3);   // last value
console.log(1 && 0 && 3);   // first falsy
console.log("" && "abc");   // first falsy
"#
        ),
        vec!["3", "0", ""]
    );
}

#[test]
fn logical_or_returns_first_truthy_or_last() {
    assert_eq!(
        run_js(
            r#"
console.log(0 || false || 3);  // first truthy
console.log(0 || false || ""); // all falsy — last value
console.log(1 || 2 || 3);      // first truthy
"#
        ),
        vec!["3", "", "1"]
    );
}

#[test]
fn nullish_coalescing_only_null_undefined() {
    assert_eq!(
        run_js(
            r#"
console.log(0 ?? "default");      // 0 — not null/undefined
console.log("" ?? "default");     // "" — not null/undefined
console.log(false ?? "default");  // false — not null/undefined
console.log(null ?? "default");   // "default"
console.log(undefined ?? "default"); // "default"
"#
        ),
        vec!["0", "", "false", "default", "default"]
    );
}

// ── comma operator ────────────────────────────────────────────────────────────

#[test]
fn comma_in_for_update() {
    assert_eq!(
        run_js(
            r#"
const result = [];
for (let i = 0, j = 10; i < 3; i++, j -= 3) {
    result.push(i + ":" + j);
}
console.log(result.join(","));
"#
        ),
        vec!["0:10,1:7,2:4"]
    );
}

// ── void operator ─────────────────────────────────────────────────────────────

#[test]
fn void_evaluates_then_returns_undefined() {
    assert_eq!(
        run_js(
            r#"
let x = 0;
const result = void (x = 42);
console.log(result);
console.log(x); // side effect happened
"#
        ),
        vec!["undefined", "42"]
    );
}

// ── typeof ────────────────────────────────────────────────────────────────────

#[test]
fn typeof_all_types() {
    assert_eq!(
        run_js(
            r#"
console.log(typeof 42);
console.log(typeof "str");
console.log(typeof true);
console.log(typeof undefined);
console.log(typeof null);
console.log(typeof {});
console.log(typeof []);
console.log(typeof function(){});
console.log(typeof Symbol());
"#
        ),
        vec![
            "number",
            "string",
            "boolean",
            "undefined",
            "object",
            "object",
            "object",
            "function",
            "symbol"
        ]
    );
}

// ── ternary chaining ──────────────────────────────────────────────────────────

#[test]
fn ternary_chain_grade() {
    assert_eq!(
        run_js(
            r#"
function grade(n) {
    return n >= 90 ? "A"
         : n >= 80 ? "B"
         : n >= 70 ? "C"
         : n >= 60 ? "D" : "F";
}
console.log(grade(95));
console.log(grade(82));
console.log(grade(55));
"#
        ),
        vec!["A", "B", "F"]
    );
}
