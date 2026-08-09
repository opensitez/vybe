/// Map and Set advanced usage — complex key types, iteration, conversion
use super::helpers::run_js;

#[test]
fn map_object_keys() {
    assert_eq!(
        run_js(
            r#"
const map = new Map();
const key1 = { id: 1 };
const key2 = { id: 2 };
map.set(key1, "value1");
map.set(key2, "value2");
console.log(map.get(key1));
console.log(map.get(key2));
console.log(map.size);
console.log(map.has({ id: 1 }));
"#
        ),
        vec!["value1", "value2", "2", "false"]
    );
}

#[test]
fn map_function_keys() {
    assert_eq!(
        run_js(
            r#"
const cache = new Map();
function memoize(fn) {
    return function(...args) {
        const key = JSON.stringify(args);
        if (!cache.has(key)) cache.set(key, fn(...args));
        return cache.get(key);
    };
}
const add = memoize((a, b) => a + b);
console.log(add(1, 2));
console.log(add(1, 2));
console.log(cache.size);
"#
        ),
        vec!["3", "3", "1"]
    );
}

#[test]
fn map_to_object_and_back() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: 1, b: 2, c: 3 };
const map = new Map(Object.entries(obj));
map.set("d", 4);
const back = Object.fromEntries(map);
console.log(back.a);
console.log(back.d);
console.log(Object.keys(back).sort().join(","));
"#
        ),
        vec!["1", "4", "a,b,c,d"]
    );
}

#[test]
fn set_operations() {
    assert_eq!(
        run_js(
            r#"
const a = new Set([1, 2, 3, 4]);
const b = new Set([3, 4, 5, 6]);
const union = new Set([...a, ...b]);
const intersection = new Set([...a].filter(x => b.has(x)));
const difference = new Set([...a].filter(x => !b.has(x)));
console.log([...union].sort((a,b)=>a-b).join(","));
console.log([...intersection].join(","));
console.log([...difference].join(","));
"#
        ),
        vec!["1,2,3,4,5,6", "3,4", "1,2"]
    );
}

#[test]
fn map_chaining_pattern() {
    assert_eq!(
        run_js(
            r#"
const freq = ["a","b","a","c","b","a"].reduce((m, v) => m.set(v, (m.get(v)??0)+1), new Map());
const sorted = [...freq.entries()].sort((a,b) => b[1]-a[1]);
console.log(sorted[0].join("="));
console.log(sorted[1].join("="));
"#
        ),
        vec!["a=3", "b=2"]
    );
}

#[test]
fn set_nan_deduplication() {
    assert_eq!(
        run_js(
            r#"
const s = new Set([NaN, NaN, 1, 1, undefined, undefined]);
console.log(s.size);
console.log(s.has(NaN));
console.log(s.has(undefined));
"#
        ),
        vec!["3", "true", "true"]
    );
}

#[test]
fn map_iteration_order() {
    assert_eq!(
        run_js(
            r#"
const m = new Map();
m.set("z", 3);
m.set("a", 1);
m.set("m", 2);
const keys = [...m.keys()];
const vals = [...m.values()];
console.log(keys.join(","));
console.log(vals.join(","));
"#
        ),
        vec!["z,a,m", "3,1,2"]
    );
}

#[test]
fn set_iteration_order_insertion() {
    assert_eq!(
        run_js(
            r#"
const s = new Set();
s.add(3); s.add(1); s.add(2); s.add(1); s.add(3);
console.log([...s].join(","));
console.log(s.size);
"#
        ),
        vec!["3,1,2", "3"]
    );
}

#[test]
fn map_as_multimap() {
    assert_eq!(
        run_js(
            r#"
class MultiMap {
    #map = new Map();
    add(key, value) {
        if (!this.#map.has(key)) this.#map.set(key, []);
        this.#map.get(key).push(value);
        return this;
    }
    get(key) { return this.#map.get(key) ?? []; }
    has(key) { return this.#map.has(key); }
}
const mm = new MultiMap();
mm.add("a", 1).add("b", 2).add("a", 3).add("a", 4);
console.log(mm.get("a").join(","));
console.log(mm.get("b").join(","));
console.log(mm.has("c"));
"#
        ),
        vec!["1,3,4", "2", "false"]
    );
}

#[test]
fn set_as_graph_adjacency() {
    assert_eq!(
        run_js(
            r#"
class Graph {
    #adj = new Map();
    addEdge(a, b) {
        if (!this.#adj.has(a)) this.#adj.set(a, new Set());
        if (!this.#adj.has(b)) this.#adj.set(b, new Set());
        this.#adj.get(a).add(b);
        this.#adj.get(b).add(a);
    }
    neighbors(node) { return [...(this.#adj.get(node) ?? [])].sort(); }
    hasEdge(a, b) { return (this.#adj.get(a) ?? new Set()).has(b); }
}
const g = new Graph();
g.addEdge("A", "B"); g.addEdge("A", "C"); g.addEdge("B", "C");
console.log(g.neighbors("A").join(","));
console.log(g.hasEdge("B", "C"));
console.log(g.hasEdge("A", "D"));
"#
        ),
        vec!["B,C", "true", "false"]
    );
}

#[test]
fn weakmap_private_data_pattern() {
    assert_eq!(
        run_js(
            r#"
const _private = new WeakMap();
class Person {
    constructor(name, age) {
        _private.set(this, { name, age });
    }
    get name() { return _private.get(this).name; }
    get age() { return _private.get(this).age; }
    birthday() { _private.get(this).age++; }
}
const p = new Person("Alice", 30);
console.log(p.name);
console.log(p.age);
p.birthday();
console.log(p.age);
console.log(_private.has(p));
"#
        ),
        vec!["Alice", "30", "31", "true"]
    );
}

#[test]
fn set_clear_returns_undefined_spec() {
    assert_eq!(
        run_js(
            r#"
const s = new Set([1, 2]);
console.log(s.clear() === undefined);
"#
        ),
        vec!["true"]
    );
}
