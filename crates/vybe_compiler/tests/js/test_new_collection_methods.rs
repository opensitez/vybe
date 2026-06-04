crate::js_cases! {
    map_groupby_groups_elements_by_string_key => {
        r#"
const entries = Map.groupBy([1, 2, 3, 4], value => value % 2 === 0 ? "even" : "odd");
console.log(entries.get("even").join(","));
console.log(entries.get("odd").join(","));
"#,
        ["2,4", "1,3"]
    };

    map_groupby_preserves_object_keys_without_string_coercion => {
        r#"
const low = { label: "low" };
const high = { label: "high" };
const grouped = Map.groupBy([1, 2, 3], value => value < 3 ? low : high);
console.log(grouped.get(low).join(","));
console.log(grouped.get(high).join(","));
console.log(grouped.has({ label: "low" }));
"#,
        ["1,2", "3", "false"]
    };

    map_groupby_tracks_number_of_groups => {
        r#"
const grouped = Map.groupBy(["ant", "bear", "cat"], value => value.length);
console.log(grouped.size);
console.log(grouped.get(3).join(","));
console.log(grouped.get(4).join(","));
"#,
        ["2", "ant,cat", "bear"]
    };

    set_union_combines_unique_values => {
        r#"
const result = new Set([1, 2, 3]).union(new Set([3, 4, 5]));
console.log([...result].join(","));
"#,
        ["1,2,3,4,5"]
    };

    set_intersection_keeps_shared_values => {
        r#"
const result = new Set([1, 2, 3, 4]).intersection(new Set([2, 4, 6]));
console.log([...result].join(","));
"#,
        ["2,4"]
    };

    set_difference_removes_rhs_values => {
        r#"
const result = new Set([1, 2, 3, 4]).difference(new Set([2, 4]));
console.log([...result].join(","));
"#,
        ["1,3"]
    };

    set_symmetric_difference_keeps_non_shared_values => {
        r#"
const result = new Set([1, 2, 3]).symmetricDifference(new Set([3, 4, 5]));
console.log([...result].join(","));
"#,
        ["1,2,4,5"]
    };

    set_issubsetof_checks_containment => {
        r#"
console.log(new Set([2, 3]).isSubsetOf(new Set([1, 2, 3, 4])));
console.log(new Set([2, 5]).isSubsetOf(new Set([1, 2, 3, 4])));
"#,
        ["true", "false"]
    };

    set_issupersetof_checks_reverse_containment => {
        r#"
console.log(new Set([1, 2, 3, 4]).isSupersetOf(new Set([2, 3])));
console.log(new Set([1, 2, 3]).isSupersetOf(new Set([2, 4])));
"#,
        ["true", "false"]
    };

    set_isdisjointfrom_checks_overlap => {
        r#"
console.log(new Set([1, 2]).isDisjointFrom(new Set([3, 4])));
console.log(new Set([1, 2]).isDisjointFrom(new Set([2, 3])));
"#,
        ["true", "false"]
    };
}
