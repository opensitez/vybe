/// Symbol usage — well-known symbols, protocols, custom protocols
use super::helpers::run_js;

#[test]
fn symbol_iterator_custom_class() {
    assert_eq!(
        run_js(
            r#"
class NumberRange {
    constructor(start, end, step = 1) {
        this.start = start; this.end = end; this.step = step;
    }
    [Symbol.iterator]() {
        let current = this.start;
        const { end, step } = this;
        return {
            next() {
                if (current <= end) { const value = current; current += step; return { value, done: false }; }
                return { value: undefined, done: true };
            }
        };
    }
}
const r = new NumberRange(1, 10, 2);
console.log([...r].join(","));
const arr2 = [...new NumberRange(10, 50, 10)];
console.log(arr2[0]);
console.log(arr2[2]);
"#
        ),
        vec!["1,3,5,7,9", "10", "30"]
    );
}

#[test]
fn symbol_has_instance() {
    assert_eq!(
        run_js(
            r#"
class EvenNumber {
    static [Symbol.hasInstance](n) {
        return typeof n === "number" && n % 2 === 0;
    }
}
console.log(2 instanceof EvenNumber);
console.log(3 instanceof EvenNumber);
console.log(100 instanceof EvenNumber);
"#
        ),
        vec!["true", "false", "true"]
    );
}

#[test]
fn symbol_to_string_tag() {
    assert_eq!(
        run_js(
            r#"
class MyCollection {
    get [Symbol.toStringTag]() { return "MyCollection"; }
}
const mc = new MyCollection();
console.log(Object.prototype.toString.call(mc));
console.log(mc.toString());
"#
        ),
        vec!["[object MyCollection]", "[object MyCollection]"]
    );
}

#[test]
fn symbol_species() {
    assert_eq!(
        run_js(
            r#"
class MyArray extends Array {
    static get [Symbol.species]() { return Array; }
    sum() { return this.reduce((a, b) => a + b, 0); }
}
const ma = new MyArray();
ma.push(1, 2, 3, 4);
const mapped = ma.map(x => x * 2);
// With Symbol.species = Array, map returns a plain Array
console.log(mapped instanceof Array);
console.log(ma instanceof MyArray);
console.log(ma.sum());
"#
        ),
        vec!["true", "true", "10"]
    );
}

#[test]
#[allow(non_snake_case)]
fn symbol_toPrimitive_all_hints() {
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
console.log(obj > 40);
"#
        ),
        vec!["42", "forty-two", "true", "true"]
    );
}

#[test]
fn symbol_as_constant_enum() {
    assert_eq!(
        run_js(
            r#"
const Direction = {
    UP: Symbol("UP"),
    DOWN: Symbol("DOWN"),
    LEFT: Symbol("LEFT"),
    RIGHT: Symbol("RIGHT") };
function move(dir) {
    switch(dir) {
        case Direction.UP: return "going up";
        case Direction.DOWN: return "going down";
        default: return "other";
    }
}
console.log(move(Direction.UP));
console.log(move(Direction.DOWN));
console.log(move(Direction.LEFT));
console.log(Direction.UP !== Direction.DOWN);
"#
        ),
        vec!["going up", "going down", "other", "true"]
    );
}

#[test]
#[allow(non_snake_case)]
fn symbol_nonEnumerable_hiding() {
    assert_eq!(
        run_js(
            r#"
const SECRET = Symbol("secret");
const obj = {
    name: "Alice",
    [SECRET]: "hidden",
    age: 30
};
console.log(Object.keys(obj).join(","));
console.log(obj[SECRET]);
console.log(Object.getOwnPropertySymbols(obj).length);
"#
        ),
        vec!["name,age", "hidden", "1"]
    );
}

#[test]
fn well_known_symbol_iterator_gen() {
    assert_eq!(
        run_js(
            r#"
class InfiniteCounter {
    constructor(start = 0) { this.n = start; }
    [Symbol.iterator]() { return this; }
    next() { return { value: this.n++, done: false }; }
}
const counter = new InfiniteCounter(5);
const first5 = [];
for (const n of counter) { first5.push(n); if (first5.length === 5) break; }
console.log(first5.join(","));
"#
        ),
        vec!["5,6,7,8,9"]
    );
}

#[test]
fn symbol_concat_spreadable() {
    assert_eq!(
        run_js(
            r#"
const arrayLike = { 0: "a", 1: "b", 2: "c", length: 3, [Symbol.isConcatSpreadable]: true };
const result = ["x"].concat(arrayLike);
console.log(result.join(","));
const notSpreadable = [1, 2];
notSpreadable[Symbol.isConcatSpreadable] = false;
const result2 = ["y"].concat(notSpreadable);
console.log(result2.length);
"#
        ),
        vec!["x,a,b,c", "2"]
    );
}

#[test]
fn global_symbol_registry() {
    assert_eq!(
        run_js(
            r#"
const s1 = Symbol.for("shared");
const s2 = Symbol.for("shared");
console.log(s1 === s2);
console.log(Symbol.keyFor(s1));
const local = Symbol("local");
console.log(Symbol.keyFor(local));
"#
        ),
        vec!["true", "shared", "undefined"]
    );
}
