/// Computed property names, shorthand, method shorthand, getter/setter shorthand
use super::helpers::run_js;

#[test]
fn computed_property_from_variable() {
    assert_eq!(
        run_js(
            r#"
const key = "name";
const obj = { [key]: "Alice" };
console.log(obj.name);
console.log(obj[key]);
"#
        ),
        vec!["Alice", "Alice"]
    );
}

#[test]
fn computed_property_from_expression() {
    assert_eq!(
        run_js(
            r#"
const prefix = "get";
const obj = {
    [`${prefix}Name`]() { return "Bob"; }
};
console.log(obj.getName());
"#
        ),
        vec!["Bob"]
    );
}

#[test]
fn shorthand_property() {
    assert_eq!(
        run_js(
            r#"
const x = 1, y = 2;
const point = { x, y };
console.log(point.x);
console.log(point.y);
"#
        ),
        vec!["1", "2"]
    );
}

#[test]
fn method_shorthand() {
    assert_eq!(
        run_js(
            r#"
const obj = {
    greet(name) { return "Hello " + name; },
    add(a, b) { return a + b; }
};
console.log(obj.greet("World"));
console.log(obj.add(3, 4));
"#
        ),
        vec!["Hello World", "7"]
    );
}

#[test]
fn async_method_shorthand() {
    assert_eq!(
        run_js(
            r#"
const obj = {
    async fetchData() {
        const v = await Promise.resolve(42);
        return v;
    }
};
async function main() {
    console.log(await obj.fetchData());
}
main();
"#
        ),
        vec!["42"]
    );
}

#[test]
fn generator_method_shorthand() {
    assert_eq!(
        run_js(
            r#"
const obj = {
    *range(n) {
        for (let i = 0; i < n; i++) yield i;
    }
};
console.log([...obj.range(4)].join(","));
"#
        ),
        vec!["0,1,2,3"]
    );
}

#[test]
fn getter_setter_shorthand_in_object() {
    assert_eq!(
        run_js(
            r#"
const obj = {
    _val: 0,
    get val() { return this._val; },
    set val(v) { this._val = v > 0 ? v : 0; }
};
obj.val = 5;
console.log(obj.val);
obj.val = -1;
console.log(obj.val);
"#
        ),
        vec!["5", "0"]
    );
}

#[test]
fn computed_symbol_key_method() {
    assert_eq!(
        run_js(
            r#"
const sym = Symbol("method");
const obj = {
    [sym]() { return "from symbol method"; }
};
console.log(obj[sym]());
"#
        ),
        vec!["from symbol method"]
    );
}

#[test]
fn property_shorthand_in_destructuring_return() {
    assert_eq!(
        run_js(
            r#"
function getPoint() {
    const x = 3, y = 4;
    return { x, y };
}
const { x, y } = getPoint();
console.log(x);
console.log(y);
"#
        ),
        vec!["3", "4"]
    );
}

#[test]
fn computed_class_method_name() {
    assert_eq!(
        run_js(
            r#"
const methodName = "doIt";
class Foo {
    [methodName]() { return 42; }
}
const f = new Foo();
console.log(f[methodName]());
console.log(f.doIt());
"#
        ),
        vec!["42", "42"]
    );
}

#[test]
fn spread_and_shorthand_combined() {
    assert_eq!(
        run_js(
            r#"
const base = { a: 1, b: 2 };
const c = 3;
const result = { ...base, c, d: 4 };
console.log(result.a);
console.log(result.c);
console.log(result.d);
"#
        ),
        vec!["1", "3", "4"]
    );
}
