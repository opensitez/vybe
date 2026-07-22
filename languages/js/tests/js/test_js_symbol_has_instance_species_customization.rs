use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `Symbol.hasInstance` & `Symbol.species` Class Hooks
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_symbol_has_instance_custom_class_check() {
    let src = r#"
class EvenNumber {
    static [Symbol.hasInstance](instance) {
        return typeof instance === "number" && instance % 2 === 0;
    }
}
console.log((2 instanceof EvenNumber) + "|" + (3 instanceof EvenNumber));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_symbol_has_instance_object_literal() {
    let src = r#"
const IntegerType = {
    [Symbol.hasInstance](val) {
        return Number.isInteger(val);
    }
};
console.log((42 instanceof IntegerType) + "|" + (3.14 instanceof IntegerType));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_symbol_has_instance_bypasses_prototype_chain() {
    let src = r#"
class Mock {}
Object.defineProperty(Mock, Symbol.hasInstance, {
    value: (inst) => inst && inst.isMock === true
});
console.log(({ isMock: true } instanceof Mock) + "|" + (new Mock() instanceof Mock));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_symbol_species_custom_array_derived_type() {
    let src = r#"
class SpecialArray extends Array {
    static get [Symbol.species]() { return Array; }
}
const sa = new SpecialArray(1, 2, 3);
const mapped = sa.map(x => x * 2);
console.log(mapped.join(",") + "|isSpecial=" + (mapped instanceof SpecialArray) + "|isArray=" + (mapped instanceof Array));
"#;
    assert_eq!(run_js(src), vec!["2,4,6|isSpecial=false|isArray=true"]);
}

#[test]
fn test_js_symbol_species_builtin_defaults_to_this() {
    let src = r#"
class DefaultSubArray extends Array {}
const dsa = new DefaultSubArray(1, 2);
const sliced = dsa.slice(0, 1);
console.log(sliced.join(",") + "|isDefaultSub=" + (sliced instanceof DefaultSubArray));
"#;
    assert_eq!(run_js(src), vec!["1|isDefaultSub=true"]);
}

#[test]
fn test_js_symbol_species_custom_promise_derivation() {
    let src = r#"
class CustomPromise extends Promise {
    static get [Symbol.species]() { return Promise; }
}
const cp = new CustomPromise(resolve => resolve("Success"));
const chained = cp.then(res => res.toUpperCase());
console.log(chained instanceof CustomPromise);
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_symbol_has_instance_non_object_rhs_throws_typeerror() {
    let src = r#"
try {
    10 instanceof null;
} catch (e) {
    console.log("instanceof Right-Hand Side Null TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["instanceof Right-Hand Side Null TypeError"]
    );
}

#[test]
fn test_js_symbol_has_instance_non_callable_throws_typeerror() {
    let src = r#"
const obj = { [Symbol.hasInstance]: "not_a_function" };
try {
    {} instanceof obj;
} catch (e) {
    console.log("Symbol.hasInstance Not Callable TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Symbol.hasInstance Not Callable TypeError"]
    );
}

#[test]
fn test_js_symbol_species_null_returns_default_base_constructor() {
    let src = r#"
class NullSpeciesArray extends Array {
    static get [Symbol.species]() { return null; }
}
const nsa = new NullSpeciesArray(10, 20);
const res = nsa.map(x => x);
console.log(res instanceof Array + "|" + (res instanceof NullSpeciesArray));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_symbol_species_getter_only() {
    let src = r#"
class ImmutableSpecies {
    static get [Symbol.species]() { return ImmutableSpecies; }
}
console.log(ImmutableSpecies[Symbol.species] === ImmutableSpecies);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_symbol_has_instance_function_prototype_default() {
    let src = r#"
function Foo() {}
console.log(Function.prototype[Symbol.hasInstance].call(Foo, new Foo()));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_symbol_species_custom_map_subclass() {
    let src = r#"
class CustomMap extends Map {
    static get [Symbol.species]() { return Map; }
}
const cm = new CustomMap([["a", 1]]);
console.log(cm instanceof CustomMap);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_symbol_has_instance_primitive_lhs_evaluation() {
    let src = r#"
class StringChecker {
    static [Symbol.hasInstance](val) {
        return typeof val === "string";
    }
}
console.log(("hello" instanceof StringChecker) + "|" + (123 instanceof StringChecker));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_symbol_species_regexp_exec_split_derivation() {
    let src = r#"
class CustomRegExp extends RegExp {
    static get [Symbol.species]() { return RegExp; }
}
const re = new CustomRegExp("a");
console.log(re.constructor[Symbol.species] === RegExp);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_symbol_has_instance_bound_function() {
    let src = r#"
function Base() {}
const BoundBase = Base.bind(null);
const inst = new Base();
console.log(inst instanceof BoundBase);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_symbol_species_typedarray_subarray_bypasses_species() {
    let src = r#"
class CustomUint8 extends Uint8Array {
    static get [Symbol.species]() { return Uint8Array; }
}
const cu8 = new CustomUint8([1, 2, 3]);
const sub = cu8.subarray(1); // TypedArray.prototype.subarray does NOT use Symbol.species!
console.log(sub instanceof CustomUint8);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_symbol_has_instance_descriptor_non_writable_non_configurable() {
    let src = r#"
const desc = Object.getOwnPropertyDescriptor(Function.prototype, Symbol.hasInstance);
console.log(desc.writable + "|" + desc.configurable);
"#;
    assert_eq!(run_js(src), vec!["false|false"]);
}

#[test]
fn test_js_symbol_species_undefined_returns_this() {
    let src = r#"
class UndefSpeciesArray extends Array {
    static get [Symbol.species]() { return undefined; }
}
const usa = new UndefSpeciesArray(1, 2);
const res = usa.map(x => x);
console.log(res instanceof Array);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_symbol_has_instance_side_effect_evaluation() {
    let src = r#"
let evaluated = false;
const trap = {
    [Symbol.hasInstance]() {
        evaluated = true;
        return true;
    }
};
console.log((100 instanceof trap) + "|Evaluated=" + evaluated);
"#;
    assert_eq!(run_js(src), vec!["true|Evaluated=true"]);
}

#[test]
fn test_js_symbol_species_well_known_symbol_identity() {
    let src = r#"
console.log(typeof Symbol.species === "symbol" && typeof Symbol.hasInstance === "symbol");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}
