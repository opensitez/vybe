//! Builtin prototype method coverage — distinct behaviors only.
crate::js_cases! {
    push_returns_new_length => {
        r#"const a=[1,2]; console.log(a.push(3,4)); console.log(a.join(","));"#,
        ["4", "1,2,3,4"]
    };

    push_empty_args_returns_same_length => {
        r#"const a=[1]; console.log(a.push()); console.log(a.length);"#,
        ["1", "1"]
    };

    push_on_empty_sets_length_one => {
        r#"const a=[]; console.log(a.push("x")); console.log(a[0]);"#,
        ["1", "x"]
    };

    pop_returns_removed_element => {
        r#"const a=[1,2,3]; console.log(a.pop()); console.log(a.join(","));"#,
        ["3", "1,2"]
    };

    pop_empty_returns_undefined => {
        r#"const a=[]; console.log(a.pop()===undefined); console.log(a.length);"#,
        ["true", "0"]
    };

    pop_single_leaves_empty => {
        r#"const a=[9]; console.log(a.pop()); console.log(a.length);"#,
        ["9", "0"]
    };

    shift_returns_first_element => {
        r#"const a=[10,20,30]; console.log(a.shift()); console.log(a.join(","));"#,
        ["10", "20,30"]
    };

    shift_empty_returns_undefined => {
        r#"const a=[]; console.log(a.shift()===undefined);"#,
        ["true"]
    };

    unshift_returns_new_length => {
        r#"const a=[2,3]; console.log(a.unshift(0,1)); console.log(a.join(","));"#,
        ["4", "0,1,2,3"]
    };

    unshift_on_empty => {
        r#"const a=[]; console.log(a.unshift(5)); console.log(a[0]);"#,
        ["1", "5"]
    };

    splice_remove_returns_removed => {
        r#"const a=[1,2,3,4]; const r=a.splice(1,2); console.log(r.join(",")); console.log(a.join(","));"#,
        ["2,3", "1,4"]
    };

    splice_insert_at_index => {
        r#"const a=[1,4]; const r=a.splice(1,0,2,3); console.log(r.join(",")); console.log(a.join(","));"#,
        ["", "1,2,3,4"]
    };

    splice_replace_elements => {
        r#"const a=[1,2,3]; const r=a.splice(1,1,"x","y"); console.log(r.join(",")); console.log(a.join(","));"#,
        ["2", "1,x,y,3"]
    };

    splice_delete_count_omitted => {
        r#"const a=[1,2,3]; const r=a.splice(1); console.log(r.join(",")); console.log(a.join(","));"#,
        ["2,3", "1"]
    };

    splice_negative_start => {
        r#"const a=[1,2,3,4]; a.splice(-2,1); console.log(a.join(","));"#,
        ["1,2,4"]
    };

    reverse_returns_same_reference => {
        r#"const a=[1,2,3]; console.log(a.reverse()===a); console.log(a.join(","));"#,
        ["true", "3,2,1"]
    };

    reverse_empty_noop => {
        r#"const a=[]; console.log(a.reverse()===a); console.log(a.length);"#,
        ["true", "0"]
    };

    sort_default_lexicographic => {
        r#"const a=[10,2,1]; a.sort(); console.log(a.join(","));"#,
        ["1,10,2"]
    };

    sort_with_numeric_comparator => {
        r#"const a=[10,2,1]; a.sort((x,y)=>x-y); console.log(a.join(","));"#,
        ["1,2,10"]
    };

    sort_returns_same_reference => {
        r#"const a=[3,1,2]; console.log(a.sort()===a);"#,
        ["true"]
    };

    sort_comparator_not_callable_throws => {
        r#"try{[1,2].sort({}); console.log("ok");}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    fill_returns_same_reference => {
        r#"const a=[1,2,3,4]; console.log(a.fill(0,1,3)===a); console.log(a.join(","));"#,
        ["true", "1,0,0,4"]
    };

    fill_without_end_fills_rest => {
        r#"const a=[1,2,3]; a.fill(9,1); console.log(a.join(","));"#,
        ["1,9,9"]
    };

    fill_entire_array => {
        r#"const a=[1,2,3]; a.fill(0); console.log(a.join(","));"#,
        ["0,0,0"]
    };

    fill_negative_start_normalized => {
        r#"const a=[1,2,3,4]; a.fill(0,-2); console.log(a.join(","));"#,
        ["1,2,0,0"]
    };

    copywithin_returns_same_reference => {
        r#"const a=[1,2,3,4,5]; console.log(a.copyWithin(0,3)===a); console.log(a.join(","));"#,
        ["true", "4,5,3,4,5"]
    };

    copywithin_partial_range => {
        r#"const a=[1,2,3,4]; a.copyWithin(1,2,3); console.log(a.join(","));"#,
        ["1,3,3,4"]
    };

    copywithin_negative_target => {
        r#"const a=[1,2,3,4]; a.copyWithin(-1,0); console.log(a.join(","));"#,
        ["1,2,3,1"]
    };

    push_spread_many_elements => {
        r#"const a=[1]; console.log(a.push(2,3,4,5)); console.log(a.length);"#,
        ["5", "5"]
    };

    pop_after_push_sequence => {
        r#"const a=[]; a.push(1); a.push(2); console.log(a.pop()); console.log(a.pop()); console.log(a.pop()===undefined);"#,
        ["2", "1", "true"]
    };

    shift_after_unshift => {
        r#"const a=[3]; a.unshift(1,2); console.log(a.shift()); console.log(a.join(","));"#,
        ["1", "2,3"]
    };

    splice_at_start => {
        r#"const a=[1,2,3]; a.splice(0,1); console.log(a.join(","));"#,
        ["2,3"]
    };

    splice_at_end => {
        r#"const a=[1,2,3]; a.splice(2,1); console.log(a.join(","));"#,
        ["1,2"]
    };

    splice_zero_delete_inserts => {
        r#"const a=[1,3]; a.splice(1,0,2); console.log(a.join(","));"#,
        ["1,2,3"]
    };

    reverse_then_pop => {
        r#"const a=[1,2,3]; a.reverse(); console.log(a.pop());"#,
        ["1"]
    };

    sort_stable_equal_elements => {
        r#"const a=[{v:1},{v:2},{v:1}]; a.sort((a,b)=>a.v-b.v); console.log(a.map(x=>x.v).join(","));"#,
        ["1,1,2"]
    };

    fill_on_sparse_skips_holes => {
        r#"const a=[1,,3]; a.fill(0); console.log(1 in a); console.log(2 in a);"#,
        ["true", "true"]
    };

    copywithin_overlapping_forward => {
        r#"const a=[1,2,3,4]; a.copyWithin(1,0,2); console.log(a.join(","));"#,
        ["1,1,2,4"]
    };

    copywithin_overlapping_backward => {
        r#"const a=[1,2,3,4]; a.copyWithin(0,1,3); console.log(a.join(","));"#,
        ["2,3,3,4"]
    };

    push_on_sparse_increases_length => {
        r#"const a=[1,,3]; console.log(a.push(4)); console.log(a.length);"#,
        ["4", "4"]
    };

    pop_does_not_shrink_below_zero => {
        r#"const a=[]; a.pop(); a.pop(); console.log(a.length);"#,
        ["0"]
    };

    unshift_multiple_on_sparse => {
        r#"const a=[,2]; console.log(a.unshift(0,1)); console.log(a.length);"#,
        ["4", "4"]
    };

    shift_on_single_element => {
        r#"const a=[99]; console.log(a.shift()); console.log(a.length);"#,
        ["99", "0"]
    };

    splice_large_delete_count => {
        r#"const a=[1,2,3,4,5]; const r=a.splice(1,10); console.log(r.length); console.log(a.join(","));"#,
        ["4", "1"]
    };

    sort_undefined_comparator_uses_default => {
        r#"const a=[3,1,2]; a.sort(undefined); console.log(a.join(","));"#,
        ["1,2,3"]
    };

    fill_start_equals_end_noop => {
        r#"const a=[1,2,3]; a.fill(9,1,1); console.log(a.join(","));"#,
        ["1,2,3"]
    };

    copywithin_end_omitted => {
        r#"const a=[1,2,3,4,5]; a.copyWithin(0,2); console.log(a.join(","));"#,
        ["3,4,5,4,5"]
    };

    // Node-verified: reverse keeps the hole a hole — `1 in a` is false.
    reverse_preserves_holes => {
        r#"const a=[1,,3]; a.reverse(); console.log(1 in a); console.log(a[0]); console.log(a[2]);"#,
        ["false", "3", "1"]
    };

    push_return_after_multiple => {
        r#"const a=[1]; const n=a.push(2,3); console.log(n); console.log(a[2]);"#,
        ["3", "3"]
    };

    unshift_return_value => {
        r#"const a=[3]; console.log(a.unshift(1,2));"#,
        ["3"]
    };

    splice_returns_empty_when_nothing_removed => {
        r#"const a=[1,2]; const r=a.splice(1,0,9); console.log(r.length); console.log(a.join(","));"#,
        ["0", "1,9,2"]
    };

    sort_empty_array => {
        r#"const a=[]; console.log(a.sort()===a); console.log(a.length);"#,
        ["true", "0"]
    };

    fill_object_reference => {
        r#"const o={}; const a=[1,2,3]; a.fill(o); console.log(a[0]===a[1]);"#,
        ["true"]
    };

    copywithin_zero_count => {
        r#"const a=[1,2,3]; a.copyWithin(1,0,0); console.log(a.join(","));"#,
        ["1,2,3"]
    };

    pop_after_splice => {
        r#"const a=[1,2,3,4]; a.splice(1,2); console.log(a.pop()); console.log(a.join(","));"#,
        ["4", "1"]
    };

    shift_after_splice => {
        r#"const a=[1,2,3,4]; a.splice(2,1); console.log(a.shift()); console.log(a.join(","));"#,
        ["1", "2,4"]
    };

    push_then_shift_fifo => {
        r#"const q=[]; q.push(1); q.push(2); console.log(q.shift()); console.log(q.shift());"#,
        ["1", "2"]
    };

    unshift_then_pop_lifo => {
        r#"const s=[]; s.unshift(1); s.unshift(2); console.log(s.pop()); console.log(s.pop());"#,
        ["1", "2"]
    };

    sort_reverse_combo => {
        r#"const a=[3,1,2]; a.sort((x,y)=>x-y); a.reverse(); console.log(a.join(","));"#,
        ["3,2,1"]
    };

    fill_negative_end => {
        r#"const a=[1,2,3,4,5]; a.fill(0,1,-1); console.log(a.join(","));"#,
        ["1,0,0,0,5"]
    };

    splice_with_negative_delete => {
        r#"const a=[1,2,3]; a.splice(1,-1,9); console.log(a.join(","));"#,
        ["1,9,2,3"]
    };

    reverse_single_element => {
        r#"const a=[42]; a.reverse(); console.log(a[0]);"#,
        ["42"]
    };

    push_on_frozen_array_throws => {
        r#"const a=Object.freeze([1]); try{a.push(2); console.log("ok");}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    sort_on_frozen_array_throws => {
        r#"const a=Object.freeze([2,1]); try{a.sort(); console.log("ok");}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    // Node-verified: fill never throws for -Infinity — relativeStart
    // clamps to 0 (§23.1.3.7), so the whole array fills.
    fill_range_error_negative_start => {
        r#"const a=[1,2,3]; a.fill(0,-Infinity); console.log(a.join(","));"#,
        ["0,0,0"]
    };

    copywithin_on_empty => {
        r#"const a=[]; console.log(a.copyWithin(0,0)===a); console.log(a.length);"#,
        ["true", "0"]
    };

    unshift_empty_args => {
        r#"const a=[1]; console.log(a.unshift()); console.log(a.length);"#,
        ["1", "1"]
    };

    // Node-verified: splice CLAMPS start to length (§23.1.3.31) — no
    // hole padding; the result is [1,9].
    splice_beyond_length_inserts_at_end => {
        r#"const a=[1]; a.splice(5,0,9); console.log(a.join(","));"#,
        ["1,9"]
    };

    pop_shortens_length_by_one => {
        r#"const a=[1,2,3]; a.pop(); console.log(a.length);"#,
        ["2"]
    };

    shift_shortens_length_by_one => {
        r#"const a=[1,2,3]; a.shift(); console.log(a.length);"#,
        ["2"]
    };

    // Node-verified: generic push on an array-like WORKS (§23.1.3.23 is
    // generic) — it sets o[0] and bumps length; no TypeError.
    push_on_array_like_object_fails => {
        r#"const o={length:0, push:Array.prototype.push}; o.push(1); console.log(o[0]);console.log(o.length);"#,
        ["1", "1"]
    };

    sort_compare_fn_returns_non_number => {
        r#"const a=[1,2]; a.sort(()=>"a"); console.log(a.length);"#,
        ["2"]
    };

    reverse_then_shift => {
        r#"const a=[1,2,3]; a.reverse(); console.log(a.shift());"#,
        ["3"]
    };

    fill_start_beyond_length_noop => {
        r#"const a=[1,2]; a.fill(0,5); console.log(a.join(","));"#,
        ["1,2"]
    };

    copywithin_start_beyond_length_noop => {
        r#"const a=[1,2]; a.copyWithin(0,5); console.log(a.join(","));"#,
        ["1,2"]
    };

    splice_delete_zero_at_end => {
        r#"const a=[1,2]; const r=a.splice(2,0,3); console.log(r.length); console.log(a.join(","));"#,
        ["0", "1,2,3"]
    };

    unshift_on_long_array => {
        r#"const a=Array.from({length:100},(_,i)=>i); a.unshift(-1); console.log(a[0]); console.log(a.length);"#,
        ["-1", "101"]
    };

    push_accepts_undefined => {
        r#"const a=[1]; a.push(undefined); console.log(a[1]===undefined);"#,
        ["true"]
    };

    // Node-verified: the hole at index 1 survives the pop — `1 in a`
    // is false.
    pop_on_sparse_array => {
        r#"const a=[1,,3]; console.log(a.pop()); console.log(1 in a);"#,
        ["3", "false"]
    };

    shift_on_sparse_array => {
        r#"const a=[,2,3]; console.log(a.shift()); console.log(a.join(","));"#,
        ["undefined", "2,3"]
    };

    // Node-verified: sort moves undefined to the END (§23.1.3.30) —
    // a[0] is 1, not undefined.
    sort_all_undefined_stable => {
        r#"const a=[undefined,1,undefined]; a.sort(); console.log(a[0]===undefined); console.log(a[2]);"#,
        ["false", "undefined"]
    };

    reverse_two_elements => {
        r#"const a=[1,2]; a.reverse(); console.log(a.join(","));"#,
        ["2,1"]
    };

    fill_with_start_only => {
        r#"const a=[1,2,3,4]; a.fill(0,2); console.log(a.join(","));"#,
        ["1,2,0,0"]
    };

    splice_returns_new_array => {
        r#"const a=[1,2,3]; const r=a.splice(0,1); console.log(Array.isArray(r)); console.log(r[0]);"#,
        ["true", "1"]
    };

    fill_then_splice_replaces_elements => {
        r#"const a=[1,2,3]; a.fill(0); a.splice(1, 1, 9); console.log(a.join(","));"#,
        ["0,9,0"]
    };

    unshift_undefined_in_array => {
        r#"const a=[1,2]; a.unshift(undefined); console.log(a.length); console.log(0 in a);"#,
        ["3", "true"]
    };

}
