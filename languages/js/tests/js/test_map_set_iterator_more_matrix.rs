crate::js_cases! {
    map_constructor_duplicate_key_last_value_wins => {
        r#"
const m = new Map([["a", 1], ["a", 2]]);
console.log(m.size);
console.log(m.get("a"));
"#,
        ["1", "2"]
    };
    map_accepts_undefined_and_null_as_distinct_keys => {
        r#"
const m = new Map([[undefined, 1], [null, 2]]);
console.log(m.size);
console.log(m.get(undefined));
console.log(m.get(null));
"#,
        ["2", "1", "2"]
    };
    map_has_object_key_after_object_mutation => {
        r#"
const key = { x: 1 };
const m = new Map([[key, 1]]);
key.x = 2;
console.log(m.has(key));
"#,
        ["true"]
    };
    map_iterator_next_after_exhaustion_has_done_true => {
        r#"
const it = new Map([["a", 1]]).keys();
it.next();
console.log(it.next().done);
"#,
        ["true"]
    };
    map_values_iterator_next_value_sequence => {
        r#"
const it = new Map([["a", 1], ["b", 2]]).values();
console.log(it.next().value);
console.log(it.next().value);
"#,
        ["1", "2"]
    };
    map_entries_iterator_next_returns_key_value_pair => {
        r#"
const pair = new Map([["a", 1]]).entries().next().value;
console.log(pair[0]);
console.log(pair[1]);
"#,
        ["a", "1"]
    };
    map_for_each_third_argument_is_map => {
        r#"
const m = new Map([["a", 1]]);
let ok = false;
m.forEach((_, __, self) => { ok = self === m; });
console.log(ok);
"#,
        ["true"]
    };
    map_size_updates_after_clear_and_reinsert => {
        r#"
const m = new Map([["a", 1], ["b", 2]]);
m.clear();
m.set("c", 3);
console.log(m.size);
console.log(m.get("c"));
"#,
        ["1", "3"]
    };
    set_accepts_null_and_undefined_as_distinct_values => {
        r#"
const s = new Set([null, undefined]);
console.log(s.size);
console.log(s.has(null));
console.log(s.has(undefined));
"#,
        ["2", "true", "true"]
    };
    set_from_string_dedupes_repeated_characters => {
        r#"
const s = new Set("banana");
console.log(Array.from(s).join(""));
"#,
        ["ban"]
    };
    set_add_same_object_twice_keeps_single_entry => {
        r#"
const obj = {};
const s = new Set([obj]);
s.add(obj);
console.log(s.size);
"#,
        ["1"]
    };
    set_iterator_next_after_exhaustion_has_done_true => {
        r#"
const it = new Set([1]).values();
it.next();
console.log(it.next().done);
"#,
        ["true"]
    };
    set_default_iterator_matches_values => {
        r#"
const a = Array.from(new Set([1, 2]).values()).join(",");
const b = Array.from(new Set([1, 2])[Symbol.iterator]()).join(",");
console.log(a === b);
"#,
        ["true"]
    };
    set_for_each_thisarg_applies_on_each_call => {
        r#"
const ctx = { base: 5 };
const out = [];
new Set([1, 2]).forEach(function(v) { out.push(v + this.base); }, ctx);
console.log(out.join(","));
"#,
        ["6,7"]
    };
    set_delete_then_add_reinserts_at_end => {
        r#"
const s = new Set([1, 2, 3]);
s.delete(2);
s.add(2);
console.log(Array.from(s).join(","));
"#,
        ["1,3,2"]
    };
    map_key_iterator_reflects_reinsertion_order => {
        r#"
const m = new Map([["a", 1], ["b", 2], ["c", 3]]);
m.delete("b");
m.set("b", 4);
console.log(Array.from(m.keys()).join(","));
"#,
        ["a,c,b"]
    };
    set_size_after_clear_and_readd_is_one => {
        r#"
const s = new Set([1, 2]);
s.clear();
s.add(3);
console.log(s.size);
"#,
        ["1"]
    };
    map_get_returns_undefined_for_deleted_key => {
        r#"
const m = new Map([["a", 1]]);
m.delete("a");
console.log(m.get("a") === undefined);
"#,
        ["true"]
    };
    set_has_returns_false_for_deleted_value => {
        r#"
const s = new Set([1, 2]);
s.delete(2);
console.log(s.has(2));
"#,
        ["false"]
    };
    map_can_store_symbol_keys => {
        r#"
const s = Symbol("k");
const m = new Map([[s, 9]]);
console.log(m.get(s));
"#,
        ["9"]
    };
    set_can_store_symbols => {
        r#"
const s = Symbol("k");
const set = new Set([s]);
console.log(set.has(s));
"#,
        ["true"]
    };
    map_can_store_bigint_keys => {
        r#"
const m = new Map([[1n, "x"]]);
console.log(m.get(1n));
"#,
        ["x"]
    };
    set_can_store_bigint_values => {
        r#"
const s = new Set([1n, 2n]);
console.log(s.has(2n));
"#,
        ["true"]
    };
    map_and_set_sizes_are_independent => {
        r#"
const m = new Map([["a", 1]]);
const s = new Set([1, 2]);
console.log(m.size + ":" + s.size);
"#,
        ["1:2"]
    };
    set_entries_iterator_done_after_all_values => {
        r#"
const it = new Set([1]).entries();
it.next();
console.log(it.next().done);
"#,
        ["true"]
    };

    map_symbol_iterator_is_entries_iterator => {
        r#"
console.log(Map.prototype[Symbol.iterator] === Map.prototype.entries);
"#,
        ["true"]
    };
}

