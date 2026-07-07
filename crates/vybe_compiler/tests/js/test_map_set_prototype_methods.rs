//! Builtin prototype method coverage — distinct behaviors only.
crate::js_cases! {
    map_set_returns_same_map => {
        r#"const m=new Map(); console.log(m.set("a",1)===m);"#,
        ["true"]
    };

    map_get_missing_undefined => {
        r#"console.log(new Map().get("x")===undefined);"#,
        ["true"]
    };

    map_has_false_for_missing => {
        r#"console.log(new Map().has("x"));"#,
        ["false"]
    };

    map_delete_returns_boolean => {
        r#"const m=new Map([["a",1]]); console.log(m.delete("a")); console.log(m.delete("a"));"#,
        ["true", "false"]
    };

    map_clear_empties => {
        r#"const m=new Map([["a",1],["b",2]]); m.clear(); console.log(m.size);"#,
        ["0"]
    };

    map_size_after_sets => {
        r#"const m=new Map(); m.set(1,1); m.set(2,2); console.log(m.size);"#,
        ["2"]
    };

    map_for_each_insertion_order => {
        r#"const m=new Map([["a",1],["b",2]]); const o=[]; m.forEach((v,k)=>o.push(k+":"+v)); console.log(o.join(","));"#,
        ["a:1,b:2"]
    };

    map_for_each_this_arg => {
        r#"const m=new Map([["a",1]]); const ctx={p:""}; m.forEach(function(v,k){this.p+=k;},ctx); console.log(ctx.p);"#,
        ["a"]
    };

    map_keys_iterator => {
        r#"console.log(Array.from(new Map([["x",1],["y",2]]).keys()).join(","));"#,
        ["x,y"]
    };

    map_values_iterator => {
        r#"console.log(Array.from(new Map([["x",1],["y",2]]).values()).join(","));"#,
        ["1,2"]
    };

    map_entries_iterator => {
        r#"const e=Array.from(new Map([["k",9]]).entries())[0]; console.log(e[0]); console.log(e[1]);"#,
        ["k", "9"]
    };

    map_set_overwrites_value => {
        r#"const m=new Map(); m.set("a",1); m.set("a",2); console.log(m.get("a")); console.log(m.size);"#,
        ["2", "1"]
    };

    map_object_key_identity => {
        r#"const k={}; const m=new Map([[k,1]]); console.log(m.get(k));"#,
        ["1"]
    };

    map_nan_key => {
        r#"const m=new Map(); m.set(NaN,1); console.log(m.get(NaN));"#,
        ["1"]
    };

    map_negative_zero_key => {
        r#"const m=new Map(); m.set(-0,"neg"); console.log(m.get(0));"#,
        ["neg"]
    };

    map_prototype_is_map => {
        r#"console.log(Map.prototype.isPrototypeOf(new Map()));"#,
        ["true"]
    };

    map_constructor_on_map => {
        r#"const m=new Map(); console.log(m.constructor===Map);"#,
        ["true"]
    };

    map_foreach_stores_undefined_value => {
        r#"const m=new Map(); m.set(undefined,"u"); console.log(m.get(undefined));"#,
        ["u"]
    };

    map_get_after_delete => {
        r#"const m=new Map([["a",1]]); m.delete("a"); console.log(m.has("a"));"#,
        ["false"]
    };

    map_chain_set_get => {
        r#"console.log(new Map().set("k",5).get("k"));"#,
        ["5"]
    };

    set_add_returns_same_set => {
        r#"const s=new Set(); console.log(s.add(1)===s);"#,
        ["true"]
    };

    set_add_duplicate_no_size_change => {
        r#"const s=new Set([1]); s.add(1); console.log(s.size);"#,
        ["1"]
    };

    set_has_true_false => {
        r#"const s=new Set([1]); console.log(s.has(1)); console.log(s.has(2));"#,
        ["true", "false"]
    };

    set_delete_returns_boolean => {
        r#"const s=new Set([1]); console.log(s.delete(1)); console.log(s.delete(1));"#,
        ["true", "false"]
    };

    set_clear_empties => {
        r#"const s=new Set([1,2]); s.clear(); console.log(s.size);"#,
        ["0"]
    };

    set_size_property => {
        r#"console.log(new Set([1,2,3]).size);"#,
        ["3"]
    };

    set_for_each_values_as_both_args => {
        r#"const s=new Set([1,2]); const o=[]; s.forEach((v,k)=>o.push(v+":"+k)); console.log(o.join(","));"#,
        ["1:1,2:2"]
    };

    set_for_each_this_arg => {
        r#"const s=new Set([1]); const ctx={n:0}; s.forEach(function(v){this.n+=v;},ctx); console.log(ctx.n);"#,
        ["1"]
    };

    set_keys_iterator => {
        r#"console.log(Array.from(new Set([3,1,2]).keys()).join(","));"#,
        ["3,1,2"]
    };

    set_values_iterator => {
        r#"console.log(Array.from(new Set(["a","b"]).values()).join(","));"#,
        ["a,b"]
    };

    set_entries_iterator => {
        r#"const p=Array.from(new Set([1]).entries())[0]; console.log(p[0]); console.log(p[1]);"#,
        ["1", "1"]
    };

    set_nan_membership => {
        r#"const s=new Set(); s.add(NaN); console.log(s.has(NaN));"#,
        ["true"]
    };

    set_negative_zero_collapses => {
        r#"const s=new Set(); s.add(-0); s.add(0); console.log(s.size);"#,
        ["1"]
    };

    set_prototype_is_set => {
        r#"console.log(Set.prototype.isPrototypeOf(new Set()));"#,
        ["true"]
    };

    set_constructor_on_set => {
        r#"const s=new Set(); console.log(s.constructor===Set);"#,
        ["true"]
    };

    set_add_object_identity => {
        r#"const o={}; const s=new Set(); s.add(o); console.log(s.has(o));"#,
        ["true"]
    };

    set_delete_missing => {
        r#"console.log(new Set().delete(1));"#,
        ["false"]
    };

    map_foreach_early_termination_not_supported => {
        r#"const m=new Map([["a",1],["b",2]]); let c=0; m.forEach(()=>c++); console.log(c);"#,
        ["2"]
    };

    set_iteration_insertion_order => {
        r#"const s=new Set([2,1,3]); console.log(Array.from(s).join(","));"#,
        ["2,1,3"]
    };

    map_from_entries_constructor => {
        r#"const m=new Map([["a",1],["b",2]]); console.log(m.size); console.log(m.get("b"));"#,
        ["2", "2"]
    };

    set_from_array_constructor => {
        r#"const s=new Set([1,1,2]); console.log(s.size);"#,
        ["2"]
    };

    map_keys_values_same_size => {
        r#"const m=new Map([["a",1],["b",2]]); console.log(Array.from(m.keys()).length); console.log(Array.from(m.values()).length);"#,
        ["2", "2"]
    };

    set_values_equals_keys => {
        r#"const s=new Set([1]); const v=Array.from(s.values())[0]; const k=Array.from(s.keys())[0]; console.log(v===k);"#,
        ["true"]
    };

    map_set_undefined_value => {
        r#"const m=new Map(); m.set("u",undefined); console.log(m.has("u")); console.log(m.get("u")===undefined);"#,
        ["true", "true"]
    };

    set_add_undefined => {
        r#"const s=new Set(); s.add(undefined); console.log(s.has(undefined));"#,
        ["true"]
    };

    map_delete_during_iterate_safe => {
        r#"const m=new Map([["a",1],["b",2]]); m.delete("a"); console.log(m.size);"#,
        ["1"]
    };

    set_add_null => {
        r#"const s=new Set(); s.add(null); console.log(s.has(null));"#,
        ["true"]
    };

    map_get_null_key => {
        r#"const m=new Map(); m.set(null,1); console.log(m.get(null));"#,
        ["1"]
    };

    map_foreach_value_only_param => {
        r#"const m=new Map([["k",7]]); let v=0; m.forEach(x=>v=x); console.log(v);"#,
        ["7"]
    };

    set_foreach_value_only_param => {
        r#"const s=new Set([9]); let v=0; s.forEach(x=>v=x); console.log(v);"#,
        ["9"]
    };

    map_prototype_has_get => {
        r#"console.log(typeof Map.prototype.get);"#,
        ["function"]
    };

    set_prototype_has_add => {
        r#"console.log(typeof Set.prototype.add);"#,
        ["function"]
    };

    map_iterator_next_done => {
        r#"const it=new Map([["a",1]]).entries(); it.next(); const r=it.next(); console.log(r.done);"#,
        ["true"]
    };

    set_iterator_next_done => {
        r#"const it=new Set([1]).values(); it.next(); const r=it.next(); console.log(r.done);"#,
        ["true"]
    };

    map_size_zero_initially => {
        r#"console.log(new Map().size);"#,
        ["0"]
    };

    set_size_zero_initially => {
        r#"console.log(new Set().size);"#,
        ["0"]
    };

    map_set_object_value => {
        r#"const o={}; const m=new Map([["k",o]]); console.log(m.get("k")===o);"#,
        ["true"]
    };

    set_add_string_primitives => {
        r#"const s=new Set(); s.add("a"); s.add("a"); console.log(s.size);"#,
        ["1"]
    };

    map_multiple_delete => {
        r#"const m=new Map([["a",1],["b",2],["c",3]]); m.delete("b"); m.delete("a"); console.log(m.size); console.log(m.has("c"));"#,
        ["1", "true"]
    };

    set_multiple_delete => {
        r#"const s=new Set([1,2,3]); s.delete(2); s.delete(1); console.log(s.size); console.log(s.has(3));"#,
        ["1", "true"]
    };

    map_for_each_third_arg_is_map => {
        r#"const m=new Map([["a",1]]); let ok=false; m.forEach((v,k,t)=>{ok=(t===m);}); console.log(ok);"#,
        ["true"]
    };

    set_for_each_third_arg_is_set => {
        r#"const s=new Set([1]); let ok=false; s.forEach((v,k,t)=>{ok=(t===s);}); console.log(ok);"#,
        ["true"]
    };

    map_entries_destructuring => {
        r#"const [[k,v]]=new Map([["x",2]]); console.log(k); console.log(v);"#,
        ["x", "2"]
    };

    set_spread_to_array => {
        r#"console.log([...new Set([1,2,2])].join(","));"#,
        ["1,2"]
    };

    map_key_string_number_distinct => {
        r#"const m=new Map(); m.set(1,"n"); m.set("1","s"); console.log(m.get(1)); console.log(m.get("1"));"#,
        ["n", "s"]
    };

    set_boolean_members => {
        r#"const s=new Set([true,false,true]); console.log(s.size); console.log(s.has(false));"#,
        ["2", "true"]
    };

    map_clear_then_set => {
        r#"const m=new Map([["a",1]]); m.clear(); m.set("b",2); console.log(m.get("b")); console.log(m.size);"#,
        ["2", "1"]
    };

    set_clear_then_add => {
        r#"const s=new Set([1]); s.clear(); s.add(2); console.log(s.has(2)); console.log(s.size);"#,
        ["true", "1"]
    };

    map_has_after_set_delete => {
        r#"const m=new Map(); m.set(1,1); m.delete(1); console.log(m.has(1));"#,
        ["false"]
    };

    set_has_after_add_delete => {
        r#"const s=new Set(); s.add(1); s.delete(1); console.log(s.has(1));"#,
        ["false"]
    };

    map_constructor_null_iterable_empty => {
        // §24.1.1.1 step 4: null (like undefined) yields an empty map — it does
        // NOT throw (node-verified). Mirrors set_constructor_undefined_iterable_ok.
        r#"try{new Map(null); console.log("ok");}catch(e){console.log(e instanceof TypeError);}"#,
        ["ok"]
    };

    set_constructor_undefined_iterable_ok => {
        r#"console.log(new Set(undefined).size);"#,
        ["0"]
    };

    map_get_returns_by_reference => {
        r#"const o={}; const m=new Map([["k",o]]); o.x=1; console.log(m.get("k").x);"#,
        ["1"]
    };

    set_values_iterator_independent => {
        r#"const s=new Set([1,2]); const a=Array.from(s.values()); s.add(3); console.log(a.length);"#,
        ["2"]
    };

    map_keys_snapshot_not_live => {
        r#"const m=new Map([["a",1]]); const k=Array.from(m.keys()); m.set("b",2); console.log(k.length);"#,
        ["1"]
    };

    set_add_returns_set_for_chain => {
        r#"console.log(new Set().add(1).add(2).size);"#,
        ["2"]
    };

    map_set_chain_three => {
        r#"const m=new Map(); m.set("a",1).set("b",2).set("c",3); console.log(m.size);"#,
        ["3"]
    };

    map_foreach_order_after_update => {
        r#"const m=new Map([["a",1]]); m.set("b",2); const o=[]; m.forEach((v,k)=>o.push(k)); console.log(o.join(","));"#,
        ["a,b"]
    };

    set_foreach_after_delete => {
        r#"const s=new Set([1,2,3]); s.delete(2); let n=0; s.forEach(()=>n++); console.log(n);"#,
        ["2"]
    };

    map_entries_value_tuple => {
        r#"const e=[...new Map([["k",1]]).entries()][0]; console.log(Array.isArray(e)); console.log(e.length);"#,
        ["true", "2"]
    };

    set_entries_value_tuple => {
        r#"const e=[...new Set([1]).entries()][0]; console.log(e[0]); console.log(e[1]);"#,
        ["1", "1"]
    };

    map_prototype_not_enumerable => {
        r#"console.log(Object.prototype.propertyIsEnumerable.call(Map.prototype,"get"));"#,
        ["false"]
    };

    set_prototype_not_enumerable => {
        r#"console.log(Object.prototype.propertyIsEnumerable.call(Set.prototype,"add"));"#,
        ["false"]
    };

    map_instance_not_plain_object => {
        r#"console.log(typeof new Map()); console.log(Array.isArray(new Map()));"#,
        ["object", "false"]
    };

    set_instance_not_plain_object => {
        r#"console.log(typeof new Set()); console.log(Array.isArray(new Set()));"#,
        ["object", "false"]
    };

    map_for_each_arrow_lexical_this => {
        r#"const m=new Map([["a",1]]); const o={v:0}; m.forEach(()=>o.v++); console.log(o.v);"#,
        ["1"]
    };

    set_for_each_arrow_lexical_this => {
        r#"const s=new Set([1,2]); const o={v:0}; s.forEach(()=>o.v++); console.log(o.v);"#,
        ["2"]
    };

}
