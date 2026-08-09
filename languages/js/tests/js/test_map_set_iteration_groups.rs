//! Map and Set iteration, grouping, and weak collection edge behaviors.

crate::js_cases! {
    map_set_get_returns_stored_value => {
        r#"const m=new Map([["a",1]]); console.log(m.get("a"));"#,
        ["1"]
    };

    map_set_overwrites_existing_key => {
        r#"const m=new Map(); m.set("k",1); m.set("k",2); console.log(m.get("k"));"#,
        ["2"]
    };

    map_has_checks_key_membership => {
        r#"const m=new Map([["x",1]]); console.log(m.has("x"));console.log(m.has("y"));"#,
        ["true", "false"]
    };

    map_delete_removes_entry => {
        r#"const m=new Map([["a",1]]); m.delete("a"); console.log(m.size);"#,
        ["0"]
    };

    map_size_reflects_entry_count => {
        r#"console.log(new Map([["a",1],["b",2]]).size);"#,
        ["2"]
    };

    map_keys_iterator_yields_keys => {
        r#"const k=new Map([["a",1]]).keys().next().value; console.log(k);"#,
        ["a"]
    };

    map_values_iterator_yields_values => {
        r#"console.log(new Map([["a",5]]).values().next().value);"#,
        ["5"]
    };

    map_entries_iterator_yields_pairs => {
        r#"const e=new Map([["k",9]]).entries().next().value; console.log(e[0]);console.log(e[1]);"#,
        ["k", "9"]
    };

    map_for_each_visits_all_entries => {
        r#"const o=[]; new Map([["a",1],["b",2]]).forEach((v,k)=>o.push(k+v)); console.log(o.sort().join(","));"#,
        ["a1,b2"]
    };

    map_object_key_uses_reference_equality => {
        r#"const k={}; const m=new Map([[k,1]]); console.log(m.get(k));console.log(m.get({}));"#,
        ["1", "undefined"]
    };

    set_number_membership_after_add => {
        r#"const s=new Set(); s.add(1); console.log(s.has(1));"#,
        ["true"]
    };

    set_add_inserts_value => {
        r#"const s=new Set([1,2]); s.delete(1); console.log(s.size);"#,
        ["1"]
    };

    set_has_checks_membership => {
        r#"console.log(new Set([1]).has(1));"#,
        ["true"]
    };

    set_size_counts_unique_values => {
        r#"console.log(new Set([1,1,2]).size);"#,
        ["2"]
    };

    set_values_iterator_same_as_keys => {
        r#"const s=new Set([1]); console.log(s.values().next().value);console.log(s.keys().next().value);"#,
        ["1", "1"]
    };

    set_entries_iterator_value_twice => {
        r#"const e=new Set([1]).entries().next().value; console.log(e[0]===e[1]);"#,
        ["true"]
    };

    set_for_each_visits_each_value => {
        r#"const o=[]; new Set([1,2]).forEach(v=>o.push(v)); console.log(o.sort().join(","));"#,
        ["1,2"]
    };

    set_object_identity_distinct => {
        r#"const o={}; const s=new Set([o,o]); console.log(s.size);"#,
        ["1"]
    };

    weakmap_set_and_get_object_key => {
        r#"const wm=new WeakMap(); const k={}; wm.set(k,1); console.log(wm.get(k));"#,
        ["1"]
    };

    weakmap_has_on_registered_key => {
        r#"const wm=new WeakMap(); const k={}; wm.set(k,1); console.log(wm.has(k));"#,
        ["true"]
    };

    weakmap_delete_removes_entry => {
        r#"const wm=new WeakMap(); const k={}; wm.set(k,1); wm.delete(k); console.log(wm.has(k));"#,
        ["false"]
    };

    weakset_add_object => {
        r#"const ws=new WeakSet(); const o={}; ws.add(o); console.log(ws.has(o));"#,
        ["true"]
    };

    weakset_delete_object => {
        r#"const ws=new WeakSet(); const o={}; ws.add(o); ws.delete(o); console.log(ws.has(o));"#,
        ["false"]
    };

    map_from_entries_constructor => {
        r#"console.log(new Map([["a",1],["b",2]]).get("b"));"#,
        ["2"]
    };

    set_from_array_constructor => {
        r#"console.log(new Set([1,2,2,3]).size);"#,
        ["3"]
    };

    map_group_by_creates_map_of_arrays => {
        r#"const g=Map.groupBy([1,2,3,4],n=>n%2===0?"even":"odd"); console.log(g.get("even").length);"#,
        ["2"]
    };

    map_iteration_order_insertion => {
        r#"const m=new Map([["b",1],["a",2]]); console.log([...m.keys()][0]);"#,
        ["b"]
    };

    set_minus_set_operation_es2025 => {
        r#"const a=new Set([1,2,3]); const b=new Set([2,3,4]); console.log(a.difference(b).size);"#,
        ["1"]
    };

    set_union_operation => {
        r#"const u=new Set([1]).union(new Set([2])); console.log(u.size);"#,
        ["2"]
    };

    set_intersection_operation => {
        r#"console.log(new Set([1,2]).intersection(new Set([2,3])).size);"#,
        ["1"]
    };

    set_symmetric_difference_operation => {
        r#"console.log(new Set([1,2]).symmetricDifference(new Set([2,3])).size);"#,
        ["2"]
    };

    set_is_subset_of => {
        r#"console.log(new Set([1,2]).isSubsetOf(new Set([1,2,3])));"#,
        ["true"]
    };

    set_is_superset_of => {
        r#"console.log(new Set([1,2,3]).isSupersetOf(new Set([1,2])));"#,
        ["true"]
    };

    set_is_disjoint_from => {
        r#"console.log(new Set([1]).isDisjointFrom(new Set([2])));"#,
        ["true"]
    };

    map_get_missing_returns_undefined => {
        r#"console.log(new Map().get("none"));"#,
        ["undefined"]
    };

    set_spread_into_array => {
        r#"console.log([...new Set([3,1,2,1])].join(","));"#,
        ["3,1,2"]
    };

    map_spread_into_array_of_pairs => {
        r#"const p=[...new Map([["a",1]])][0]; console.log(p[0]);console.log(p[1]);"#,
        ["a", "1"]
    };

    map_chain_set_get => {
        r#"console.log(new Map().set("x",1).get("x"));"#,
        ["1"]
    };

    weakmap_not_iterable_no_size => {
        r#"const wm=new WeakMap(); wm.set({},"a"); console.log(typeof wm.size);"#,
        ["undefined"]
    };

    map_symbol_key => {
        r#"const s=Symbol("k"); const m=new Map([[s,1]]); console.log(m.get(s));"#,
        ["1"]
    };

    set_delete_missing_returns_false => {
        r#"console.log(new Set([1]).delete(2));"#,
        ["false"]
    };

    map_delete_missing_returns_false => {
        r#"console.log(new Map().delete("x"));"#,
        ["false"]
    };

    set_symbol_iterator_is_values_iterator => {
        r#"console.log(Set.prototype[Symbol.iterator] === Set.prototype.values);"#,
        ["true"]
    };
}
