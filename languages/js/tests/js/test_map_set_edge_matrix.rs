crate::js_cases! {
    map_constructor_from_iterable_sets_entries => {
        r#"
const m = new Map([["a", 1], ["b", 2]]);
console.log(m.size);
console.log(m.get("a"));
console.log(m.get("b"));
"#,
        ["2", "1", "2"]
    };

    map_constructor_empty_has_zero_size => {
        r#"
const m = new Map();
console.log(m.size);
"#,
        ["0"]
    };

    map_get_missing_returns_undefined => {
        r#"
const m = new Map();
console.log(m.get("missing") === undefined);
"#,
        ["true"]
    };

    map_set_returns_same_map_for_chaining => {
        r#"
const m = new Map();
console.log(m.set("a", 1) === m);
"#,
        ["true"]
    };

    map_object_keys_use_identity => {
        r#"
const a = {};
const b = {};
const m = new Map([[a, 1], [b, 2]]);
console.log(m.get(a));
console.log(m.get(b));
"#,
        ["1", "2"]
    };

    map_nan_keys_compare_equal => {
        r#"
const m = new Map();
m.set(NaN, "x");
console.log(m.get(NaN));
console.log(m.has(NaN));
"#,
        ["x", "true"]
    };

    map_negative_zero_and_positive_zero_are_same_key => {
        r#"
const m = new Map();
m.set(-0, "neg");
console.log(m.get(0));
console.log(m.size);
"#,
        ["neg", "1"]
    };

    map_delete_missing_returns_false => {
        r#"
const m = new Map();
console.log(m.delete("x"));
"#,
        ["false"]
    };

    map_delete_present_returns_true => {
        r#"
const m = new Map([["x", 1]]);
console.log(m.delete("x"));
console.log(m.has("x"));
"#,
        ["true", "false"]
    };

    map_clear_returns_undefined => {
        r#"
const m = new Map([["x", 1]]);
console.log(m.clear() === undefined);
console.log(m.size);
"#,
        ["true", "0"]
    };

    map_for_each_visits_entries_in_insertion_order => {
        r#"
const m = new Map([["a", 1], ["b", 2], ["c", 3]]);
const out = [];
m.forEach((value, key) => out.push(key + ":" + value));
console.log(out.join(","));
"#,
        ["a:1,b:2,c:3"]
    };

    map_for_each_receives_this_arg => {
        r#"
const m = new Map([["a", 1], ["b", 2]]);
const ctx = { prefix: ">" };
const out = [];
m.forEach(function(value, key) { out.push(this.prefix + key + value); }, ctx);
console.log(out.join(","));
"#,
        [">a1,>b2"]
    };

    map_keys_iterator_preserves_order => {
        r#"
const m = new Map([["a", 1], ["b", 2]]);
console.log(Array.from(m.keys()).join(","));
"#,
        ["a,b"]
    };

    map_values_iterator_preserves_order => {
        r#"
const m = new Map([["a", 1], ["b", 2]]);
console.log(Array.from(m.values()).join(","));
"#,
        ["1,2"]
    };

    map_entries_iterator_preserves_order => {
        r#"
const m = new Map([["a", 1], ["b", 2]]);
const out = [];
for (const [k, v] of m.entries()) out.push(k + ":" + v);
console.log(out.join(","));
"#,
        ["a:1,b:2"]
    };

    map_default_iterator_matches_entries => {
        r#"
const m = new Map([["a", 1], ["b", 2]]);
const out = [];
for (const [k, v] of m) out.push(k + ":" + v);
console.log(out.join(","));
"#,
        ["a:1,b:2"]
    };

    map_reset_existing_key_preserves_original_position => {
        r#"
const m = new Map([["a", 1], ["b", 2]]);
m.set("a", 9);
console.log(Array.from(m.keys()).join(","));
console.log(m.get("a"));
"#,
        ["a,b", "9"]
    };

    map_delete_then_readd_moves_key_to_end => {
        r#"
const m = new Map([["a", 1], ["b", 2]]);
m.delete("a");
m.set("a", 3);
console.log(Array.from(m.keys()).join(","));
"#,
        ["b,a"]
    };

    map_size_counts_distinct_object_keys => {
        r#"
const m = new Map();
m.set({}, 1);
m.set({}, 2);
console.log(m.size);
"#,
        ["2"]
    };

    set_constructor_from_iterable_dedupes_values => {
        r#"
const s = new Set([1, 2, 2, 3]);
console.log(s.size);
"#,
        ["3"]
    };

    set_add_returns_same_set_for_chaining => {
        r#"
const s = new Set();
console.log(s.add(1) === s);
"#,
        ["true"]
    };

    set_has_missing_returns_false => {
        r#"
const s = new Set([1, 2]);
console.log(s.has(3));
"#,
        ["false"]
    };

    set_delete_missing_returns_false => {
        r#"
const s = new Set([1, 2]);
console.log(s.delete(3));
"#,
        ["false"]
    };

    set_delete_present_returns_true => {
        r#"
const s = new Set([1, 2]);
console.log(s.delete(2));
console.log(s.has(2));
"#,
        ["true", "false"]
    };

    set_clear_returns_undefined => {
        r#"
const s = new Set([1, 2]);
console.log(s.clear() === undefined);
console.log(s.size);
"#,
        ["true", "0"]
    };

    set_nan_values_dedupe => {
        r#"
const s = new Set([NaN, NaN]);
console.log(s.size);
console.log(s.has(NaN));
"#,
        ["1", "true"]
    };

    set_negative_zero_and_positive_zero_are_same_value => {
        r#"
const s = new Set();
s.add(-0);
console.log(s.has(0));
console.log(s.size);
"#,
        ["true", "1"]
    };

    set_object_identity_keeps_distinct_objects => {
        r#"
const s = new Set();
s.add({});
s.add({});
console.log(s.size);
"#,
        ["2"]
    };

    set_iteration_preserves_insertion_order => {
        r#"
const s = new Set(["a", "b", "c"]);
console.log(Array.from(s).join(","));
"#,
        ["a,b,c"]
    };

    set_keys_matches_values_iterator => {
        r#"
const s = new Set([1, 2]);
console.log(Array.from(s.keys()).join(","));
console.log(Array.from(s.values()).join(","));
"#,
        ["1,2", "1,2"]
    };

    set_entries_yields_value_value_pairs => {
        r#"
const s = new Set([1, 2]);
const out = [];
for (const [a, b] of s.entries()) out.push(a + ":" + b);
console.log(out.join(","));
"#,
        ["1:1,2:2"]
    };

    set_for_each_receives_value_value_set => {
        r#"
const s = new Set([1, 2]);
const out = [];
s.forEach((value, key, set) => out.push([value, key, set === s].join(":")));
console.log(out.join(","));
"#,
        ["1:1:true,2:2:true"]
    };

    set_for_each_honors_this_arg => {
        r#"
const s = new Set([1, 2]);
const ctx = { mul: 10 };
const out = [];
s.forEach(function(value) { out.push(value * this.mul); }, ctx);
console.log(out.join(","));
"#,
        ["10,20"]
    };

    set_delete_then_readd_moves_value_to_end => {
        r#"
const s = new Set(["a", "b"]);
s.delete("a");
s.add("a");
console.log(Array.from(s).join(","));
"#,
        ["b,a"]
    };

    map_string_and_number_keys_are_distinct => {
        r#"
const m = new Map();
m.set("1", "string");
m.set(1, "number");
console.log(m.size);
console.log(m.get("1"));
console.log(m.get(1));
"#,
        ["2", "string", "number"]
    };

    map_foreach_mutation_during_iteration => {
        r#"
const m = new Map([["a", 1]]);
const out = [];
m.forEach((v, k) => {
    out.push(k);
    if (k === "a") m.set("b", 2);
});
console.log(out.join(","));
"#,
        ["a,b"]
    };
}
