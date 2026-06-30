//! Array copy, splice, slice, and search method behaviors.

crate::js_cases! {
    array_slice_returns_shallow_copy => {
        r#"const a=[1,2,3]; const s=a.slice(); s[0]=9; console.log(a[0]);"#,
        ["1"]
    };

    array_slice_with_start_index => {
        r#"console.log([1,2,3,4].slice(2).join(","));"#,
        ["3,4"]
    };

    array_slice_with_start_and_end => {
        r#"console.log([1,2,3,4].slice(1,3).join(","));"#,
        ["2,3"]
    };

    array_slice_negative_start => {
        r#"console.log([1,2,3].slice(-2).join(","));"#,
        ["2,3"]
    };

    array_splice_remove_insert => {
        r#"const a=[1,2,3,4]; a.splice(1,2,9,8); console.log(a.join(","));"#,
        ["1,9,8,4"]
    };

    array_splice_returns_removed => {
        r#"console.log([1,2,3].splice(1,1)[0]);"#,
        ["2"]
    };

    array_splice_at_end_appends => {
        r#"const a=[1,2]; a.splice(2,0,3); console.log(a.join(","));"#,
        ["1,2,3"]
    };

    array_copywithin_copies_range => {
        r#"const a=[1,2,3,4]; a.copyWithin(0,2); console.log(a.join(","));"#,
        ["3,4,3,4"]
    };

    array_copywithin_with_end => {
        r#"const a=[1,2,3,4,5]; a.copyWithin(0,3,4); console.log(a.join(","));"#,
        ["4,2,3,4,5"]
    };

    array_concat_flattens_one_level => {
        r#"console.log([1].concat([2,3]).join(","));"#,
        ["1,2,3"]
    };

    array_concat_does_not_mutate => {
        r#"const a=[1]; a.concat(2); console.log(a.length);"#,
        ["1"]
    };

    array_indexof_finds_first => {
        r#"console.log([1,2,2,3].indexOf(2));"#,
        ["1"]
    };

    array_lastindexof_finds_last => {
        r#"console.log([1,2,2,3].lastIndexOf(2));"#,
        ["2"]
    };

    array_includes_finds_value => {
        r#"console.log([1,2,3].includes(2));"#,
        ["true"]
    };

    array_includes_nan_false => {
        r#"console.log([NaN].includes(NaN));"#,
        ["false"]
    };

    array_find_returns_first_match => {
        r#"console.log([1,4,3].find(x=>x>2));"#,
        ["4"]
    };

    array_find_returns_undefined_when_missing => {
        r#"console.log([1,2].find(x=>x>5));"#,
        ["undefined"]
    };

    array_findindex_returns_index => {
        r#"console.log([1,4,3].findIndex(x=>x>2));"#,
        ["1"]
    };

    array_findindex_minus_one_when_missing => {
        r#"console.log([1,2].findIndex(x=>x>9));"#,
        ["-1"]
    };

    array_findlast_es2023 => {
        r#"console.log([1,4,3,4].findLast(x=>x>2));"#,
        ["4"]
    };

    array_findlastindex_es2023 => {
        r#"console.log([1,4,3,4].findLastIndex(x=>x>2));"#,
        ["3"]
    };

    array_filter_creates_new_array => {
        r#"const a=[1,2,3]; const f=a.filter(x=>x>1); console.log(f.join(","));console.log(a.length);"#,
        ["2,3", "3"]
    };

    array_map_transforms_elements => {
        r#"console.log([1,2,3].map(x=>x*2).join(","));"#,
        ["2,4,6"]
    };

    array_reduce_left_fold => {
        r#"console.log([1,2,3].reduce((a,b)=>a+b,0));"#,
        ["6"]
    };

    array_reduceright_fold => {
        r#"console.log([1,2,3].reduceRight((a,b)=>a-b,0));"#,
        ["0"]
    };

    array_some_short_circuits => {
        r#"console.log([1,2,3].some(x=>x===2));"#,
        ["true"]
    };

    array_every_all_match => {
        r#"console.log([2,4,6].every(x=>x%2===0));"#,
        ["true"]
    };

    array_flatten_one_level => {
        r#"console.log([1,[2,3]].flat().join(","));"#,
        ["1,2,3"]
    };

    array_flatten_depth_two => {
        r#"console.log([1,[2,[3]]].flat(2).join(","));"#,
        ["1,2,3"]
    };

    array_flatmap_maps_and_flattens => {
        r#"console.log([1,2].flatMap(x=>[x,x]).join(","));"#,
        ["1,1,2,2"]
    };

    array_join_default_comma => {
        r#"console.log([1,2,3].join());"#,
        ["1,2,3"]
    };

    array_reverse_mutates => {
        r#"const a=[1,2,3]; a.reverse(); console.log(a[0]);"#,
        ["3"]
    };

    array_sort_default_string_order => {
        r#"console.log([3,1,2].sort().join(","));"#,
        ["1,2,3"]
    };

    array_fill_sets_range => {
        r#"console.log([0,0,0].fill(7,1,3).join(","));"#,
        ["0,7,7"]
    };

    array_at_negative_index => {
        r#"console.log([1,2,3].at(-1));"#,
        ["3"]
    };

    array_toreversed_non_mutating => {
        r#"const a=[1,2]; const r=a.toReversed(); console.log(a[0]);console.log(r[0]);"#,
        ["1", "2"]
    };

    array_tosorted_non_mutating => {
        r#"const a=[3,1]; const s=a.toSorted(); console.log(a[0]);console.log(s[0]);"#,
        ["3", "1"]
    };

    array_tospliced_non_mutating => {
        r#"const a=[1,2,3]; const s=a.toSpliced(1,1,9); console.log(a[1]);console.log(s[1]);"#,
        ["2", "9"]
    };

    array_with_non_mutating_index_set => {
        r#"const a=[1,2,3]; const w=a.with(1,9); console.log(a[1]);console.log(w[1]);"#,
        ["2", "9"]
    };

    array_indexof_from_index => {
        r#"console.log([1,2,1].indexOf(1,1));"#,
        ["2"]
    };

    array_includes_from_index => {
        r#"console.log([1,2,3].includes(1,1));"#,
        ["false"]
    };

    array_splice_delete_count_zero_inserts => {
        r#"const a=[1,2]; a.splice(1,0,9); console.log(a.join(","));"#,
        ["1,9,2"]
    };

    array_slice_on_sparse_preserves_holes => {
        r#"const a=[1,,3]; console.log(a.slice().length);"#,
        ["3"]
    };

    array_concat_spreads_strings => {
        r#"console.log([1].concat("ab").join(","));"#,
        ["1,a,b"]
    };

    array_find_on_empty_undefined => {
        r#"console.log([].find(()=>true));"#,
        ["undefined"]
    };

    array_reduce_on_single_element => {
        r#"console.log([5].reduce((a,b)=>a+b));"#,
        ["5"]
    };

    array_map_skips_holes => {
        r#"let n=0; [1,,3].map(()=>n++); console.log(n);"#,
        ["2"]
    };

    array_filter_skips_holes => {
        r#"console.log([1,,3].filter(x=>x>0).length);"#,
        ["2"]
    };

    array_includes_searches_nan_with_samevaluezero => {
        r#"console.log([NaN].indexOf(NaN));"#,
        ["-1"]
    };
}
