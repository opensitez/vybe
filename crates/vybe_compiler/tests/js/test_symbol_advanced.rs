/// Symbol advanced — Symbol.for/keyFor global registry, well-known symbols
/// (toPrimitive, iterator, hasInstance, species, toStringTag, isConcatSpreadable),
/// Symbol as private-like key, Symbol description.
use super::helpers::run_js;

#[test]
fn symbol_uniqueness() {
    assert_eq!(
        run_js(
            r#"
const a = Symbol("x");
const b = Symbol("x");
console.log(a === b);
console.log(typeof a);
"#
        ),
        vec!["false", "symbol"]
    );
}

#[test]
fn symbol_for_registry_returns_same() {
    assert_eq!(
        run_js(
            r#"
const a = Symbol.for("shared");
const b = Symbol.for("shared");
console.log(a === b);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn symbol_keyfor_returns_key() {
    assert_eq!(
        run_js(
            r#"
const sym = Symbol.for("myKey");
console.log(Symbol.keyFor(sym));
const local = Symbol("local");
console.log(Symbol.keyFor(local));
"#
        ),
        vec!["myKey", "undefined"]
    );
}

#[test]
fn symbol_description_property() {
    assert_eq!(
        run_js(
            r#"
const s = Symbol("hello");
console.log(s.description);
const anon = Symbol();
console.log(anon.description);
"#
        ),
        vec!["hello", "undefined"]
    );
}

#[test]
fn symbol_not_in_for_in() {
    assert_eq!(
        run_js(
            r#"
const sym = Symbol("hidden");
const obj = { visible: 1, [sym]: 2 };
const keys = [];
for (const k in obj) keys.push(k);
console.log(keys.join(","));
console.log(keys.includes("hidden"));
"#
        ),
        vec!["visible", "false"]
    );
}

#[test]
fn symbol_not_in_json() {
    assert_eq!(
        run_js(
            r#"
const sym = Symbol("s");
const obj = { [sym]: 42, normal: 1 };
const json = JSON.stringify(obj);
console.log(json);
"#
        ),
        vec!["{\"normal\":1}"]
    );
}

#[test]
fn symbol_as_private_like_key() {
    assert_eq!(
        run_js(
            r#"
const _private = Symbol("private");
function makeObj(secret) {
    const obj = {};
    obj[_private] = secret;
    obj.getSecret = function() { return this[_private]; };
    return obj;
}
const o = makeObj("shhh");
console.log(o.getSecret());
console.log(o[_private]);
console.log(Object.keys(o).join(","));
"#
        ),
        vec!["shhh", "shhh", "getSecret"]
    );
}

#[test]
fn symbol_iterator_makes_object_iterable() {
    assert_eq!(
        run_js(
            r#"
const range = {
    from: 1, to: 5,
    [Symbol.iterator]() {
        let cur = this.from;
        const to = this.to;
        return {
            next() {
                return cur <= to
                    ? { value: cur++, done: false }
                    : { done: true };
            }
        };
    }
};
console.log([...range].join(","));
"#
        ),
        vec!["1,2,3,4,5"]
    );
}

#[test]
fn symbol_to_string_tag_custom() {
    assert_eq!(
        run_js(
            r#"
class MyCollection {
    get [Symbol.toStringTag]() { return "MyCollection"; }
}
const c = new MyCollection();
console.log(Object.prototype.toString.call(c));
"#
        ),
        vec!["[object MyCollection]"]
    );
}

#[test]
fn symbol_has_instance_custom() {
    assert_eq!(
        run_js(
            r#"
class OddNumbers {
    static [Symbol.hasInstance](n) {
        return typeof n === "number" && n % 2 !== 0;
    }
}
console.log(1 instanceof OddNumbers);
console.log(2 instanceof OddNumbers);
console.log(3 instanceof OddNumbers);
"#
        ),
        vec!["true", "false", "true"]
    );
}

#[test]
fn symbol_is_concat_spreadable() {
    assert_eq!(
        run_js(
            r#"
const arrayLike = {
    length: 2, 0: "a", 1: "b",
    [Symbol.isConcatSpreadable]: true
};
const result = [].concat(arrayLike);
console.log(result.join(","));
"#
        ),
        vec!["a,b"]
    );
}

#[test]
fn symbol_species_in_subclass() {
    assert_eq!(
        run_js(
            r#"
class MyArray extends Array {
    static get [Symbol.species]() { return Array; }
}
const m = new MyArray(1, 2, 3);
const mapped = m.map(x => x * 2);
console.log(mapped instanceof Array);
console.log(mapped instanceof MyArray); // false due to Symbol.species
"#
        ),
        vec!["true", "false"]
    );
}

#[test]
fn symbol_to_primitive_all_hints() {
    assert_eq!(
        run_js(
            r#"
const obj = {
    [Symbol.toPrimitive](hint) {
        if (hint === "number") return 42;
        if (hint === "string") return "forty-two";
        return true;
    }
};
console.log(+obj);
console.log(`${obj}`);
console.log(obj + "");
"#
        ),
        vec!["42", "forty-two", "true"]
    );
}

#[test]
fn symbol_split_custom() {
    assert_eq!(
        run_js(
            r#"
class CaseInsensitiveSplit {
    constructor(sep) { this.sep = sep.toLowerCase(); }
    [Symbol.split](str) {
        return str.toLowerCase().split(this.sep);
    }
}
const result = "Hello-WORLD-foo".split(new CaseInsensitiveSplit("-"));
console.log(result.join(","));
"#
        ),
        vec!["hello,world,foo"]
    );
}
