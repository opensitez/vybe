/// Final coverage — edge cases not covered elsewhere

use super::helpers::run_js;

#[test]
fn comma_operator_in_for() {
    assert_eq!(run_js(r#"
let a = 0, b = 10;
for (let i = 0; i < 5; i++, b--) { a += i; }
console.log(a);
console.log(b);
"#), vec!["10", "5"]);
}

#[test]
fn short_circuit_assignment() {
    assert_eq!(run_js(r#"
let count = 0;
const inc = () => ++count;
false && inc();
true || inc();
null ?? inc();
console.log(count);
true && inc();
false || inc();
"something" ?? inc();
console.log(count);
"#), vec!["1", "3"]);
}

#[test]
fn object_shorthand_method_name() {
    assert_eq!(run_js(r#"
const name = "greet";
const obj = {
    [name]() { return "hello"; },
    get value() { return 42; },
};
console.log(obj.greet());
console.log(obj.value);
console.log(typeof obj.greet);
"#), vec!["hello", "42", "function"]);
}

#[test]
fn string_raw_tag() {
    assert_eq!(run_js(r#"
console.log(String.raw`\n\t\r`);
console.log(String.raw`Hello\nWorld`.length);
"#), vec!["\\n\\t\\r", "12"]);
}

#[test]
fn nullish_assign_short_circuit() {
    assert_eq!(run_js(r#"
let calls = 0;
const expensive = () => ++calls;
let a = "existing";
a ??= expensive();
console.log(a);
console.log(calls);
let b = null;
b ??= expensive();
console.log(b);
console.log(calls);
"#), vec!["existing", "0", "1", "1"]);
}

#[test]
fn logical_or_assign_semantics() {
    assert_eq!(run_js(r#"
let calls = 0;
const fn = () => ++calls;
let a = "truthy";
a ||= fn();
console.log(a);
console.log(calls);
let b = 0;
b ||= fn();
console.log(b);
console.log(calls);
"#), vec!["truthy", "0", "1", "1"]);
}

#[test]
fn logical_and_assign_semantics() {
    assert_eq!(run_js(r#"
let calls = 0;
const fn = () => ++calls;
let a = 0;
a &&= fn();
console.log(a);
console.log(calls);
let b = "truthy";
b &&= fn();
console.log(b);
console.log(calls);
"#), vec!["0", "0", "1", "1"]);
}

#[test]
fn array_destructure_iterator_once() {
    assert_eq!(run_js(r#"
let iterCount = 0;
const iterable = {
    [Symbol.iterator]() {
        iterCount++;
        let i = 0;
        return { next() { return i < 3 ? { value: i++, done: false } : { done: true }; } };
    }
};
const [a, b, c] = iterable;
console.log(a);
console.log(c);
console.log(iterCount);
"#), vec!["0", "2", "1"]);
}
