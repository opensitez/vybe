use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Arrow Functions (`() => {}`), Lexical `this`, `arguments`, `super`, `new.target`
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_arrow_function_lexical_this_enclosing_scope() {
    let src = r#"
const obj = {
    name: "LexicalObj",
    getArrow() {
        return () => this.name;
    }
};
const fn = obj.getArrow();
console.log(fn());
"#;
    assert_eq!(run_js(src), vec!["LexicalObj"]);
}

#[test]
fn test_js_arrow_function_lexical_this_detached_invocation() {
    let src = r#"
const obj = {
    count: 10,
    incrementLater() {
        const fn = () => ++this.count;
        return fn;
    }
};
const inc = obj.incrementLater();
console.log(inc() + "|" + inc());
"#;
    assert_eq!(run_js(src), vec!["11|12"]);
}

#[test]
fn test_js_arrow_function_no_own_arguments_object() {
    let src = r#"
function outer(a, b) {
    const arrow = () => arguments[0] + arguments[1]; // Refers to outer function's arguments!
    return arrow();
}
console.log(outer(10, 20));
"#;
    assert_eq!(run_js(src), vec!["30"]);
}

#[test]
fn test_js_arrow_function_in_global_scope_arguments_throws_referenceerror() {
    let src = r#"
const arrow = () => typeof arguments;
console.log(arrow());
"#;
    assert_eq!(run_js(src), vec!["undefined"]);
}

#[test]
fn test_js_arrow_function_cannot_be_constructed_with_new_throws_typeerror() {
    let src = r#"
const Arrow = () => {};
try {
    new Arrow();
} catch (e) {
    console.log("Arrow Constructor TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Arrow Constructor TypeError"]);
}

#[test]
fn test_js_arrow_function_has_no_prototype_property() {
    let src = r#"
const arrow = () => {};
console.log(arrow.prototype === undefined);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_arrow_function_lexical_super_property_access() {
    let src = r#"
class Base {
    greet() { return "BaseGreet"; }
}
class Derived extends Base {
    getArrowGreet() {
        return () => super.greet();
    }
}
const d = new Derived();
console.log(d.getArrowGreet()());
"#;
    assert_eq!(run_js(src), vec!["BaseGreet"]);
}

#[test]
fn test_js_arrow_function_lexical_new_target() {
    let src = r#"
function Base() {
    const getNewTarget = () => new.target;
    return getNewTarget();
}
console.log(new Base() === Base);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_arrow_function_concise_body_implicit_return() {
    let src = r#"
const square = x => x * x;
console.log(square(5));
"#;
    assert_eq!(run_js(src), vec!["25"]);
}

#[test]
fn test_js_arrow_function_concise_body_object_literal_parentheses() {
    let src = r#"
const makeObj = (k, v) => ({ [k]: v });
console.log(makeObj("a", 100).a);
"#;
    assert_eq!(run_js(src), vec!["100"]);
}

#[test]
fn test_js_arrow_function_block_body_explicit_return() {
    let src = r#"
const add = (a, b) => {
    const sum = a + b;
    return sum;
};
console.log(add(3, 4));
"#;
    assert_eq!(run_js(src), vec!["7"]);
}

#[test]
fn test_js_arrow_function_rest_parameters() {
    let src = r#"
const sumAll = (...nums) => nums.reduce((acc, n) => acc + n, 0);
console.log(sumAll(1, 2, 3, 4));
"#;
    assert_eq!(run_js(src), vec!["10"]);
}

#[test]
fn test_js_arrow_function_destructured_parameters() {
    let src = r#"
const getFullName = ({ first, last }) => `${first} ${last}`;
console.log(getFullName({ first: "John", last: "Doe" }));
"#;
    assert_eq!(run_js(src), vec!["John Doe"]);
}

#[test]
fn test_js_arrow_function_default_parameter_values() {
    let src = r#"
const greet = (name = "Guest") => `Hello ${name}`;
console.log(greet() + "|" + greet("Alice"));
"#;
    assert_eq!(run_js(src), vec!["Hello Guest|Hello Alice"]);
}

#[test]
fn test_js_nested_arrow_functions_lexical_this() {
    let src = r#"
const obj = {
    val: 99,
    getDeep() {
        return () => () => () => this.val;
    }
};
console.log(obj.getDeep()()()());
"#;
    assert_eq!(run_js(src), vec!["99"]);
}

#[test]
fn test_js_arrow_function_async_syntax() {
    let src = r#"
const asyncFetch = async (val) => await Promise.resolve("Fetched: " + val);
(async () => {
    console.log(await asyncFetch("Data"));
})();
"#;
    assert_eq!(run_js(src), vec!["Fetched: Data"]);
}

#[test]
fn test_js_arrow_function_generator_syntax_prohibited() {
    let src = r#"
try {
    eval("const gen = *() => {};");
} catch (e) {
    console.log("Arrow Generator SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Arrow Generator SyntaxError"]);
}

#[test]
fn test_js_arrow_function_name_inference() {
    let src = r#"
const myArrow = () => {};
console.log(myArrow.name);
"#;
    assert_eq!(run_js(src), vec!["myArrow"]);
}

#[test]
fn test_js_arrow_function_length_property() {
    let src = r#"
const fn = (a, b = 1, c) => {};
console.log(fn.length); // length counts parameters before first default parameter!
"#;
    assert_eq!(run_js(src), vec!["1"]);
}

#[test]
fn test_js_arrow_function_in_array_higher_order_methods() {
    let src = r#"
const arr = [1, 2, 3];
const doubled = arr.map(x => x * 2);
console.log(doubled.join(","));
"#;
    assert_eq!(run_js(src), vec!["2,4,6"]);
}

#[test]
fn test_js_arrow_function_lexical_this_derived_constructor_after_super() {
    let src = r#"
class Base { constructor() { this.x = 10; } }
class Derived extends Base {
    constructor() {
        super();
        const getX = () => this.x;
        this.res = getX();
    }
}
console.log(new Derived().res);
"#;
    assert_eq!(run_js(src), vec!["10"]);
}
