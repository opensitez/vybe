/// Data transformation patterns — reduce to object, groupBy, index, pivot

use super::helpers::run_js;

#[test]
fn reduce_to_object_by_key() {
    assert_eq!(run_js(r#"
const users = [
    { id: 1, name: "Alice" },
    { id: 2, name: "Bob" },
    { id: 3, name: "Charlie" }
];
const indexed = users.reduce((acc, u) => ({ ...acc, [u.id]: u }), {});
console.log(indexed[1].name);
console.log(indexed[3].name);
"#), vec!["Alice", "Charlie"]);
}

#[test]
fn group_by_property() {
    assert_eq!(run_js(r#"
const items = [
    { cat: "A", val: 1 }, { cat: "B", val: 2 },
    { cat: "A", val: 3 }, { cat: "B", val: 4 }, { cat: "C", val: 5 }
];
const grouped = items.reduce((acc, item) => {
    const key = item.cat;
    (acc[key] ??= []).push(item.val);
    return acc;
}, {});
console.log(grouped.A.join(","));
console.log(grouped.B.join(","));
console.log(grouped.C.join(","));
"#), vec!["1,3", "2,4", "5"]);
}

#[test]
fn count_occurrences_reduce() {
    assert_eq!(run_js(r#"
const words = ["apple", "banana", "apple", "cherry", "banana", "apple"];
const counts = words.reduce((acc, w) => {
    acc[w] = (acc[w] ?? 0) + 1;
    return acc;
}, {});
console.log(counts.apple);
console.log(counts.banana);
console.log(counts.cherry);
"#), vec!["3", "2", "1"]);
}

#[test]
fn flatten_and_deduplicate() {
    assert_eq!(run_js(r#"
const nested = [[1, 2, 3], [2, 3, 4], [4, 5]];
const unique = [...new Set(nested.flat())].sort((a, b) => a - b);
console.log(unique.join(","));
"#), vec!["1,2,3,4,5"]);
}

#[test]
fn pivot_table_pattern() {
    assert_eq!(run_js(r#"
const sales = [
    { region: "North", product: "A", amount: 100 },
    { region: "South", product: "A", amount: 200 },
    { region: "North", product: "B", amount: 150 },
    { region: "South", product: "B", amount: 250 },
];
const pivot = sales.reduce((acc, s) => {
    if (!acc[s.region]) acc[s.region] = {};
    acc[s.region][s.product] = (acc[s.region][s.product] ?? 0) + s.amount;
    return acc;
}, {});
console.log(pivot.North.A);
console.log(pivot.South.B);
"#), vec!["100", "250"]);
}

#[test]
fn transform_keys() {
    assert_eq!(run_js(r#"
function mapKeys(obj, fn) {
    return Object.fromEntries(
        Object.entries(obj).map(([k, v]) => [fn(k), v])
    );
}
const obj = { firstName: "Alice", lastName: "Smith" };
const snaked = mapKeys(obj, k => k.replace(/([A-Z])/g, '_$1').toLowerCase());
console.log(snaked.first_name);
console.log(snaked.last_name);
"#), vec!["Alice", "Smith"]);
}

#[test]
fn merge_arrays_by_key() {
    assert_eq!(run_js(r#"
function mergeBy(key, ...arrays) {
    const map = new Map();
    for (const arr of arrays) {
        for (const item of arr) {
            const k = item[key];
            map.set(k, { ...(map.get(k) ?? {}), ...item });
        }
    }
    return [...map.values()];
}
const names = [{ id: 1, name: "Alice" }, { id: 2, name: "Bob" }];
const ages = [{ id: 1, age: 30 }, { id: 2, age: 25 }];
const merged = mergeBy("id", names, ages);
merged.sort((a, b) => a.id - b.id);
console.log(merged[0].name);
console.log(merged[0].age);
console.log(merged[1].name);
"#), vec!["Alice", "30", "Bob"]);
}

#[test]
fn deep_sum_of_nested() {
    assert_eq!(run_js(r#"
function deepSum(obj) {
    if (typeof obj === "number") return obj;
    if (Array.isArray(obj)) return obj.reduce((s, x) => s + deepSum(x), 0);
    if (typeof obj === "object") return Object.values(obj).reduce((s, v) => s + deepSum(v), 0);
    return 0;
}
const data = { a: 1, b: [2, 3, { c: 4 }], d: { e: 5, f: [6, 7] } };
console.log(deepSum(data));
"#), vec!["28"]);
}

#[test]
fn sort_by_multiple_keys() {
    assert_eq!(run_js(r#"
const people = [
    { name: "Bob", age: 30 },
    { name: "Alice", age: 25 },
    { name: "Charlie", age: 30 },
    { name: "Alice", age: 20 },
];
people.sort((a, b) => {
    if (a.name !== b.name) return a.name.localeCompare(b.name);
    return a.age - b.age;
});
console.log(people.map(p => p.name + p.age).join(","));
"#), vec!["Alice20,Alice25,Bob30,Charlie30"]);
}
