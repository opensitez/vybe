/// Object.groupBy / Map.groupBy (ES2024), array grouping patterns,
/// grouping with complex keys, empty groups, nested grouping.

use super::helpers::run_js;

// ── Object.groupBy ────────────────────────────────────────────────────────────

#[test]
fn object_groupby_basic() {
    assert_eq!(run_js(r#"
const items = [1, 2, 3, 4, 5, 6];
const groups = Object.groupBy(items, n => n % 2 === 0 ? "even" : "odd");
console.log(groups.even.join(","));
console.log(groups.odd.join(","));
"#), vec!["2,4,6", "1,3,5"]);
}

#[test]
fn object_groupby_by_string_property() {
    assert_eq!(run_js(r#"
const people = [
    { name: "Alice", dept: "eng" },
    { name: "Bob", dept: "eng" },
    { name: "Carol", dept: "hr" },
    { name: "Dave", dept: "hr" }
];
const groups = Object.groupBy(people, p => p.dept);
console.log(groups.eng.length);
console.log(groups.hr.length);
console.log(groups.eng[0].name);
"#), vec!["2", "2", "Alice"]);
}

#[test]
fn object_groupby_returns_null_prototype_object() {
    assert_eq!(run_js(r#"
const groups = Object.groupBy([1, 2, 3], n => "key");
console.log("key" in groups);
console.log(groups.key.length);
"#), vec!["true", "3"]);
}

#[test]
fn object_groupby_preserves_order_within_groups() {
    assert_eq!(run_js(r#"
const letters = ["c", "a", "b", "a", "c", "b"];
const groups = Object.groupBy(letters, l => l);
console.log(groups.a.join(","));
console.log(groups.b.join(","));
console.log(groups.c.join(","));
"#), vec!["a,a", "b,b", "c,c"]);
}

#[test]
fn object_groupby_single_element_groups() {
    assert_eq!(run_js(r#"
const words = ["apple", "banana", "cherry"];
const groups = Object.groupBy(words, w => w[0]);
console.log(groups.a[0]);
console.log(groups.b[0]);
console.log(groups.c[0]);
"#), vec!["apple", "banana", "cherry"]);
}

#[test]
fn object_groupby_empty_array() {
    assert_eq!(run_js(r#"
const groups = Object.groupBy([], x => x);
console.log(Object.keys(groups).length);
"#), vec!["0"]);
}

#[test]
fn object_groupby_number_keys() {
    assert_eq!(run_js(r#"
const nums = [10, 21, 35, 42, 57];
const groups = Object.groupBy(nums, n => Math.floor(n / 10) * 10);
const keys = Object.keys(groups).sort((a, b) => +a - +b);
console.log(keys.join(","));
"#), vec!["10,20,30,40,50"]);
}

// ── Map.groupBy ───────────────────────────────────────────────────────────────

#[test]
fn map_groupby_basic() {
    assert_eq!(run_js(r#"
const items = [1, 2, 3, 4, 5];
const groups = Map.groupBy(items, n => n % 2 === 0 ? "even" : "odd");
console.log(groups instanceof Map);
console.log(groups.get("even").join(","));
console.log(groups.get("odd").join(","));
"#), vec!["true", "2,4", "1,3,5"]);
}

#[test]
fn map_groupby_with_object_keys() {
    assert_eq!(run_js(r#"
const keyA = { type: "A" };
const keyB = { type: "B" };
const items = [
    { val: 1, key: keyA },
    { val: 2, key: keyB },
    { val: 3, key: keyA }
];
const groups = Map.groupBy(items, item => item.key);
console.log(groups.get(keyA).length);
console.log(groups.get(keyB).length);
"#), vec!["2", "1"]);
}

#[test]
fn map_groupby_preserves_key_identity() {
    assert_eq!(run_js(r#"
const groups = Map.groupBy([1, 2, 3], n => n > 1 ? "big" : "small");
console.log(groups.size);
console.log(groups.has("big"));
console.log(groups.has("small"));
"#), vec!["2", "true", "true"]);
}

// ── grouping with complex logic ───────────────────────────────────────────────

#[test]
fn groupby_range_buckets() {
    assert_eq!(run_js(r#"
const scores = [45, 67, 89, 92, 78, 55, 33];
const grades = Object.groupBy(scores, s => {
    if (s >= 90) return "A";
    if (s >= 70) return "B";
    if (s >= 50) return "C";
    return "F";
});
console.log(grades.A.join(","));
console.log(grades.B.join(","));
console.log(grades.F.join(","));
"#), vec!["92", "89,78", "45,33"]);
}

#[test]
fn groupby_then_map_groups() {
    assert_eq!(run_js(r#"
const data = [1, 2, 3, 4, 5, 6];
const groups = Object.groupBy(data, n => n % 2 === 0 ? "even" : "odd");
const sums = {};
for (const [key, vals] of Object.entries(groups)) {
    sums[key] = vals.reduce((a, b) => a + b, 0);
}
console.log(sums.even);
console.log(sums.odd);
"#), vec!["12", "9"]);
}

// ── nested grouping ───────────────────────────────────────────────────────────

#[test]
fn nested_groupby_two_levels() {
    assert_eq!(run_js(r#"
const data = [
    { dept: "eng", level: "senior" },
    { dept: "eng", level: "junior" },
    { dept: "hr", level: "senior" }
];
const byDept = Object.groupBy(data, d => d.dept);
const engByLevel = Object.groupBy(byDept.eng, d => d.level);
console.log(engByLevel.senior.length);
console.log(engByLevel.junior.length);
"#), vec!["1", "1"]);
}
