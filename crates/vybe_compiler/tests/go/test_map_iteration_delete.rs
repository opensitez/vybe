//! Map iteration with delete, `clear`, nil-map read/write compile, and two-value `range`.
//!
//! Distinct from `test_maps_patterns_extra.rs` (general map CRUD) and
//! `test_range_iteration_extra.rs` (slice/string/channel range). Focuses on
//! deleting entries while ranging, draining maps, `clear`, and nil-map edge cases.
//!
//! Runtime `clear(m)` cases document a known VM gap (`clear` compiles but is not
//! callable at runtime); see also `test_slice_copy_clear.rs`.


go_run_cases! {
    map_delete_all_keys_during_single_value_range =>
        ("package main; import \"fmt\"; func main() { values := map[string]int{\"a\": 1, \"b\": 2, \"c\": 3}; for key := range values { delete(values, key) }; fmt.Println(len(values)) }", vec!["0"]),

    map_delete_matching_value_during_two_value_range =>
        ("package main; import \"fmt\"; func main() { values := map[string]int{\"a\": 1, \"b\": 2, \"c\": 2}; for key, value := range values { if value == 2 { delete(values, key) } }; fmt.Println(len(values)); fmt.Println(values[\"a\"]) }", vec!["1", "1"]),

    map_delete_values_above_threshold_during_range =>
        ("package main; import \"fmt\"; func main() { values := map[string]int{\"low\": 1, \"mid\": 5, \"high\": 9}; for key, value := range values { if value > 5 { delete(values, key) } }; fmt.Println(len(values)); total := 0; for _, value := range values { total += value }; fmt.Println(total) }", vec!["2", "6"]),

    map_delete_one_key_preserves_remaining_lookups =>
        ("package main; import \"fmt\"; func main() { values := map[string]int{\"keep\": 10, \"drop\": 20, \"stay\": 30}; delete(values, \"drop\"); fmt.Println(len(values)); fmt.Println(values[\"keep\"]); fmt.Println(values[\"stay\"]) }", vec!["2", "10", "30"]),

    map_delete_even_values_sum_remaining =>
        ("package main; import \"fmt\"; func main() { values := map[string]int{\"a\": 1, \"b\": 2, \"c\": 3, \"d\": 4}; for key, value := range values { if value%2 == 0 { delete(values, key) } }; total := 0; for _, value := range values { total += value }; fmt.Println(total) }", vec!["4"]),

    map_delete_none_when_condition_always_false =>
        ("package main; import \"fmt\"; func main() { values := map[string]int{\"x\": 1, \"y\": 2}; for key, value := range values { if value < 0 { delete(values, key) } }; fmt.Println(len(values)) }", vec!["2"]),

    map_clear_builtin_drains_three_entries =>
        ("package main; import \"fmt\"; func main() { values := map[string]int{\"a\": 1, \"b\": 2, \"c\": 3}; clear(values); fmt.Println(len(values)) }", vec!["0"]),

    map_clear_builtin_then_reinsert_one =>
        ("package main; import \"fmt\"; func main() { values := map[string]int{\"old\": 9}; clear(values); values[\"new\"] = 4; fmt.Println(len(values)); fmt.Println(values[\"new\"]) }", vec!["1", "4"]),

    map_clear_on_already_empty_map =>
        ("package main; import \"fmt\"; func main() { values := map[int]int{}; clear(values); fmt.Println(len(values)) }", vec!["0"]),

    map_clear_on_nil_map_is_noop =>
        ("package main; import \"fmt\"; func main() { var values map[string]int; clear(values); fmt.Println(values == nil); fmt.Println(len(values)) }", vec!["true", "0"]),

    nil_map_two_value_range_zero_iterations =>
        ("package main; import \"fmt\"; func main() { var values map[string]int; count := 0; for key, value := range values { _, _ = key, value; count++ }; fmt.Println(count) }", vec!["0"]),

    map_two_value_range_sum_all_values =>
        ("package main; import \"fmt\"; func main() { values := map[string]int{\"a\": 10, \"b\": 20, \"c\": 30}; total := 0; for _, value := range values { total += value }; fmt.Println(total) }", vec!["60"]),

    map_two_value_range_count_keys_via_len =>
        ("package main; import \"fmt\"; func main() { values := map[string]int{\"x\": 1, \"y\": 2, \"z\": 3}; count := 0; for key, _ := range values { _ = key; count++ }; fmt.Println(count) }", vec!["3"]),

    map_two_value_range_blank_key_value_sum =>
        ("package main; import \"fmt\"; func main() { values := map[int]int{1: 4, 2: 5, 3: 6}; total := 0; for _, value := range values { total += value }; fmt.Println(total) }", vec!["15"]),

    map_two_value_range_int_keys_lookup_sum =>
        ("package main; import \"fmt\"; func main() { values := map[int]string{1: \"a\", 2: \"b\"}; fmt.Println(len(values[1])); fmt.Println(len(values[2])) }", vec!["1", "1"]),

    map_delete_during_range_leaves_single_entry =>
        ("package main; import \"fmt\"; func main() { values := map[string]int{\"only\": 7, \"gone\": 1}; for key, value := range values { if value == 1 { delete(values, key) } }; fmt.Println(len(values)); fmt.Println(values[\"only\"]) }", vec!["1", "7"]),

    map_delete_all_during_range_stays_empty =>
        ("package main; import \"fmt\"; func main() { values := map[string]int{\"a\": 1, \"b\": 2}; for key := range values { delete(values, key) }; fmt.Println(len(values)) }", vec!["0"]),

    map_delete_named_keys_during_range =>
        ("package main; import \"fmt\"; func main() { values := map[string]int{\"pre_a\": 1, \"pre_b\": 2, \"ok\": 3}; for key := range values { if key == \"pre_a\" || key == \"pre_b\" { delete(values, key) } }; fmt.Println(len(values)); fmt.Println(values[\"ok\"]) }", vec!["1", "3"]),

    nil_map_read_missing_key_returns_zero =>
        ("package main; import \"fmt\"; func main() { var values map[string]int; fmt.Println(values[\"absent\"]) }", vec!["0"]),

    nil_map_delete_stays_nil =>
        ("package main; import \"fmt\"; func main() { var values map[string]int; delete(values, \"x\"); fmt.Println(values == nil) }", vec!["true"]),

    map_two_value_range_bool_key_values =>
        ("package main; import \"fmt\"; func main() { values := map[bool]int{true: 5, false: 7}; total := 0; for _, value := range values { total += value }; fmt.Println(total) }", vec!["12"]),

    map_delete_reinsert_under_new_key_during_range =>
        ("package main; import \"fmt\"; func main() { values := map[string]int{\"old\": 0}; for key, value := range values { if value == 0 { delete(values, key); values[\"fresh\"] = 9 } }; fmt.Println(len(values)); fmt.Println(values[\"fresh\"]) }", vec!["1", "9"]),
}

go_compile_cases! {
    nil_map_index_write_compile =>
        "package main; func main() { var values map[string]int; values[\"k\"] = 1 }",

    nil_map_read_in_expression_compile =>
        "package main; func main() { var values map[string]int; _ = values[\"k\"] + 1 }",

    nil_map_two_value_range_with_body_compile =>
        "package main; func main() { var values map[string]int; for key, value := range values { _, _ = key, value } }",

    nil_map_clear_compile =>
        "package main; func main() { var values map[string]int; clear(values) }",

    delete_current_key_during_two_value_range_compile =>
        "package main; func main() { values := map[string]int{\"a\": 1}; for key, value := range values { if value == 1 { delete(values, key) } } }",

    delete_during_nested_map_of_slice_range_compile =>
        "package main; func main() { outer := map[string]map[string]int{\"x\": {\"a\": 1}}; for _, inner := range outer { for key := range inner { delete(inner, key) } } }",

    map_two_value_range_assign_outer_vars_compile =>
        "package main; func main() { values := map[string]int{\"a\": 1}; var key string; var value int; for key, value = range values { _, _ = key, value } }",

    delete_in_range_with_continue_compile =>
        "package main; func main() { values := map[int]int{1: 1, 2: 2}; for key, value := range values { if value == 1 { delete(values, key); continue }; _ = key } }",

    clear_then_two_value_range_compile =>
        "package main; func main() { values := map[string]int{\"a\": 1}; clear(values); for key, value := range values { _, _ = key, value } }",

    nil_map_delete_then_range_compile =>
        "package main; func main() { var values map[string]int; delete(values, \"z\"); for key := range values { _ = key } }",

    map_range_key_value_in_switch_compile =>
        "package main; func main() { values := map[string]int{\"a\": 1}; for key, value := range values { switch key { case \"a\": _ = value } } }",

    delete_during_range_labeled_break_compile =>
        "package main; func main() { values := map[string]int{\"a\": 1, \"b\": 2}; outer: for key := range values { delete(values, key); break outer } }",
}
