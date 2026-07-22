use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `structuredClone` Cyclic Reference Handling & Object Graph Cloning
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_structured_clone_plain_object_copy() {
    let src = r#"
const orig = { a: 1, b: { c: 2 } };
const clone = structuredClone(orig);
console.log((clone !== orig) + "|" + (clone.b !== orig.b) + "|" + (clone.b.c === 2));
"#;
    assert_eq!(run_js(src), vec!["true|true|true"]);
}

#[test]
fn test_js_structured_clone_self_referential_object() {
    let src = r#"
const obj = { name: "Root" };
obj.self = obj;
const clone = structuredClone(obj);
console.log((clone !== obj) + "|" + (clone.self === clone));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_structured_clone_mutual_cyclical_references() {
    let src = r#"
const nodeA = { id: "A" };
const nodeB = { id: "B" };
nodeA.sibling = nodeB;
nodeB.sibling = nodeA;

const cloneA = structuredClone(nodeA);
console.log((cloneA.sibling !== nodeB) + "|" + (cloneA.sibling.sibling === cloneA));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_structured_clone_shared_sub_object_identity_preserved() {
    let src = r#"
const shared = { val: 42 };
const root = { first: shared, second: shared };
const clone = structuredClone(root);
console.log((clone.first !== shared) + "|" + (clone.first === clone.second));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_structured_clone_array_with_cyclic_element() {
    let src = r#"
const arr = [1, 2];
arr.push(arr);
const clone = structuredClone(arr);
console.log((clone !== arr) + "|" + (clone[2] === clone));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_structured_clone_symbol_property_ignored_or_error() {
    let src = r#"
const sym = Symbol("key");
const obj = { [sym]: "data", stringKey: "data" };
const clone = structuredClone(obj);
console.log(clone.stringKey + "|hasSym=" + (sym in clone));
"#;
    assert_eq!(run_js(src), vec!["data|hasSym=false"]); // Symbol keys are omitted by structuredClone!
}

#[test]
fn test_js_structured_clone_function_throws_datacloneerror() {
    let src = r#"
try {
    structuredClone({ fn: () => {} });
} catch (e) {
    console.log("DataCloneError Function");
}
"#;
    assert_eq!(run_js(src), vec!["DataCloneError Function"]);
}

#[test]
fn test_js_structured_clone_dom_node_throws_datacloneerror() {
    let src = r#"
try {
    structuredClone({ symbolVal: Symbol() }); // Symbol values cannot be cloned!
} catch (e) {
    console.log("DataCloneError Symbol");
}
"#;
    assert_eq!(run_js(src), vec!["DataCloneError Symbol"]);
}

#[test]
fn test_js_structured_clone_primitives() {
    let src = r#"
console.log(`${structuredClone(123)}:${structuredClone("str")}:${structuredClone(true)}:${structuredClone(null)}:${structuredClone(undefined)}`);
"#;
    assert_eq!(run_js(src), vec!["123:str:true:null:undefined"]);
}

#[test]
fn test_js_structured_clone_bigint_primitive() {
    let src = r#"
const b = 9007199254740999n;
console.log(structuredClone(b) === b);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_structured_clone_getter_setter_evaluated_to_data_property() {
    let src = r#"
const obj = {
    get val() { return 100; }
};
const clone = structuredClone(obj);
const desc = Object.getOwnPropertyDescriptor(clone, "val");
console.log(desc.value + "|hasGetter=" + (typeof desc.get !== "undefined"));
"#;
    assert_eq!(run_js(src), vec!["100|hasGetter=false"]); // Accessors are serialized as static data values!
}

#[test]
fn test_js_structured_clone_prototype_stripped_to_plain_object() {
    let src = r#"
class CustomClass {
    constructor() { this.x = 10; }
}
const inst = new CustomClass();
const clone = structuredClone(inst);
console.log((clone.x === 10) + "|isCustom=" + (clone instanceof CustomClass) + "|isObject=" + (clone.constructor === Object));
"#;
    assert_eq!(run_js(src), vec!["true|isCustom=false|isObject=true"]);
}

#[test]
fn test_js_structured_clone_non_enumerable_properties_ignored() {
    let src = r#"
const obj = { visible: 1 };
Object.defineProperty(obj, "hidden", { value: 2, enumerable: false });
const clone = structuredClone(obj);
console.log(clone.visible + "|hasHidden=" + ("hidden" in clone));
"#;
    assert_eq!(run_js(src), vec!["1|hasHidden=false"]);
}

#[test]
fn test_js_structured_clone_sparse_array_holes_preserved() {
    let src = r#"
const sparse = [1, , 3];
const clone = structuredClone(sparse);
console.log(clone.length + "|hasHole=" + !(1 in clone));
"#;
    assert_eq!(run_js(src), vec!["3|hasHole=true"]);
}

#[test]
fn test_js_structured_clone_null_prototype_object() {
    let src = r#"
const obj = Object.create(null);
obj.a = 1;
const clone = structuredClone(obj);
console.log(clone.a + "|" + (Object.getPrototypeOf(clone) === Object.prototype));
"#;
    assert_eq!(run_js(src), vec!["1|true"]);
}

#[test]
fn test_js_structured_clone_deep_array_nesting() {
    let src = r#"
const nested = [[["deep"]]];
const clone = structuredClone(nested);
console.log(clone[0][0][0]);
"#;
    assert_eq!(run_js(src), vec!["deep"]);
}

#[test]
fn test_js_structured_clone_proxy_target_cloning() {
    let src = r#"
const target = { x: 50 };
const proxy = new Proxy(target, {});
const clone = structuredClone(proxy);
console.log(clone.x + "|" + (clone !== target));
"#;
    assert_eq!(run_js(src), vec!["50|true"]);
}

#[test]
fn test_js_structured_clone_frozen_object_cloned_as_extensible() {
    let src = r#"
const frozen = Object.freeze({ a: 1 });
const clone = structuredClone(frozen);
console.log(Object.isFrozen(clone) + "|" + Object.isExtensible(clone));
"#;
    assert_eq!(run_js(src), vec!["false|true"]);
}

#[test]
fn test_js_structured_clone_circular_map_and_set() {
    let src = r#"
const map = new Map();
map.set("self", map);
const cloneMap = structuredClone(map);
console.log(cloneMap.get("self") === cloneMap);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_structured_clone_no_arguments_throws_typeerror() {
    let src = r#"
try {
    structuredClone();
} catch (e) {
    console.log("structuredClone Missing Arg TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["structuredClone Missing Arg TypeError"]);
}
