//! Go 1.21+ `maps` Keys, Values, Equal, and EqualFunc — distinct from
//! `test_slices_maps_stdlib.rs` (Clone/Copy/DeleteFunc) and map iteration in
//! `test_map_iteration_delete.rs`.

use crate::helpers::*;

go_run_cases! {
    maps_keys_int_map_len => (
        "package main; import \"fmt\"; import \"maps\"; func main() { m := map[int]string{1: \"a\", 2: \"b\", 3: \"c\"}; keys := maps.Keys(m); fmt.Println(len(keys)) }",
        vec!["3"]
    ),
    maps_keys_string_map_collect => (
        "package main; import \"fmt\"; import \"maps\"; func main() { m := map[string]int{\"x\": 1, \"y\": 2}; keys := maps.Keys(m); found := 0; for _, k := range keys { if k == \"x\" || k == \"y\" { found++ } }; fmt.Println(found) }",
        vec!["2"]
    ),
    maps_keys_empty_map => (
        "package main; import \"fmt\"; import \"maps\"; func main() { keys := maps.Keys(map[string]int{}); fmt.Println(len(keys)) }",
        vec!["0"]
    ),
    maps_keys_nil_map => (
        "package main; import \"fmt\"; import \"maps\"; func main() { var m map[int]bool; keys := maps.Keys(m); fmt.Println(keys == nil); fmt.Println(len(keys)) }",
        vec!["true", "0"]
    ),
    maps_keys_single_entry => (
        "package main; import \"fmt\"; import \"maps\"; func main() { m := map[int]int{42: 99}; keys := maps.Keys(m); fmt.Println(len(keys)); fmt.Println(keys[0]) }",
        vec!["1", "42"]
    ),
    maps_keys_bool_key_type => (
        "package main; import \"fmt\"; import \"maps\"; func main() { m := map[bool]string{true: \"yes\", false: \"no\"}; fmt.Println(len(maps.Keys(m))) }",
        vec!["2"]
    ),
    maps_keys_struct_value_type => (
        "package main; import \"fmt\"; import \"maps\"; type Pt struct { X int }; func main() { m := map[string]Pt{\"a\": {1}, \"b\": {2}}; fmt.Println(len(maps.Keys(m))) }",
        vec!["2"]
    ),
    maps_values_int_map_len => (
        "package main; import \"fmt\"; import \"maps\"; func main() { m := map[string]int{\"a\": 10, \"b\": 20, \"c\": 30}; vals := maps.Values(m); fmt.Println(len(vals)) }",
        vec!["3"]
    ),
    maps_values_sum_iteration => (
        "package main; import \"fmt\"; import \"maps\"; func main() { m := map[int]int{1: 10, 2: 20, 3: 30}; sum := 0; for v := range maps.Values(m) { sum += v }; fmt.Println(sum) }",
        vec!["60"]
    ),
    maps_values_empty_map => (
        "package main; import \"fmt\"; import \"maps\"; func main() { vals := maps.Values(map[int]string{}); fmt.Println(len(vals)) }",
        vec!["0"]
    ),
    maps_values_nil_map => (
        "package main; import \"fmt\"; import \"maps\"; func main() { var m map[string]int; vals := maps.Values(m); fmt.Println(vals == nil); fmt.Println(len(vals)) }",
        vec!["true", "0"]
    ),
    maps_values_string_elements => (
        "package main; import \"fmt\"; import \"maps\"; func main() { m := map[int]string{1: \"go\", 2: \"vybe\"}; count := 0; for v := range maps.Values(m) { if len(v) >= 2 { count++ } }; fmt.Println(count) }",
        vec!["2"]
    ),
    maps_values_bool_values => (
        "package main; import \"fmt\"; import \"maps\"; func main() { m := map[string]bool{\"a\": true, \"b\": false}; trues := 0; for v := range maps.Values(m) { if v { trues++ } }; fmt.Println(trues) }",
        vec!["1"]
    ),
    maps_equal_identical_maps => (
        "package main; import \"fmt\"; import \"maps\"; func main() { a := map[string]int{\"x\": 1, \"y\": 2}; b := map[string]int{\"x\": 1, \"y\": 2}; fmt.Println(maps.Equal(a, b)) }",
        vec!["true"]
    ),
    maps_equal_different_values => (
        "package main; import \"fmt\"; import \"maps\"; func main() { a := map[int]string{1: \"a\"}; b := map[int]string{1: \"b\"}; fmt.Println(maps.Equal(a, b)) }",
        vec!["false"]
    ),
    maps_equal_different_keys => (
        "package main; import \"fmt\"; import \"maps\"; func main() { a := map[int]int{1: 10}; b := map[int]int{2: 10}; fmt.Println(maps.Equal(a, b)) }",
        vec!["false"]
    ),
    maps_equal_different_lengths => (
        "package main; import \"fmt\"; import \"maps\"; func main() { a := map[int]int{1: 1, 2: 2}; b := map[int]int{1: 1}; fmt.Println(maps.Equal(a, b)) }",
        vec!["false"]
    ),
    maps_equal_both_empty => (
        "package main; import \"fmt\"; import \"maps\"; func main() { fmt.Println(maps.Equal(map[string]int{}, map[string]int{})) }",
        vec!["true"]
    ),
    maps_equal_both_nil => (
        "package main; import \"fmt\"; import \"maps\"; func main() { var a, b map[int]int; fmt.Println(maps.Equal(a, b)) }",
        vec!["true"]
    ),
    maps_equal_nil_vs_empty => (
        "package main; import \"fmt\"; import \"maps\"; func main() { var a map[int]int; b := map[int]int{}; fmt.Println(maps.Equal(a, b)) }",
        vec!["true"]
    ),
    maps_equal_one_nil_other_nonempty => (
        "package main; import \"fmt\"; import \"maps\"; func main() { var a map[int]int; b := map[int]int{1: 1}; fmt.Println(maps.Equal(a, b)) }",
        vec!["false"]
    ),
    maps_equal_same_key_different_value => (
        "package main; import \"fmt\"; import \"maps\"; func main() { a := map[string]int{\"k\": 5}; b := map[string]int{\"k\": 6}; fmt.Println(maps.Equal(a, b)) }",
        vec!["false"]
    ),
    maps_equal_func_custom_value_compare_true => (
        "package main; import \"fmt\"; import \"maps\"; func main() { a := map[int]int{1: 10, 2: 20}; b := map[int]int{1: 10, 2: 20}; fmt.Println(maps.EqualFunc(a, b, func(x, y int) bool { return x == y })) }",
        vec!["true"]
    ),
    maps_equal_func_custom_compare_false => (
        "package main; import \"fmt\"; import \"maps\"; func main() { a := map[int]int{1: 10}; b := map[int]int{1: 11}; fmt.Println(maps.EqualFunc(a, b, func(x, y int) bool { return x == y })) }",
        vec!["false"]
    ),
    maps_equal_func_abs_value_equivalence => (
        "package main; import \"fmt\"; import \"maps\"; func main() { a := map[int]int{1: -5}; b := map[int]int{1: 5}; eq := maps.EqualFunc(a, b, func(x, y int) bool { if x < 0 { x = -x }; if y < 0 { y = -y }; return x == y }); fmt.Println(eq) }",
        vec!["true"]
    ),
    maps_equal_func_different_key_sets => (
        "package main; import \"fmt\"; import \"maps\"; func main() { a := map[int]int{1: 1}; b := map[int]int{2: 1}; fmt.Println(maps.EqualFunc(a, b, func(x, y int) bool { return x == y })) }",
        vec!["false"]
    ),
    maps_equal_func_both_empty => (
        "package main; import \"fmt\"; import \"maps\"; func main() { fmt.Println(maps.EqualFunc(map[int]int{}, map[int]int{}, func(x, y int) bool { return x == y })) }",
        vec!["true"]
    ),
    maps_equal_func_both_nil => (
        "package main; import \"fmt\"; import \"maps\"; func main() { var a, b map[string]int; fmt.Println(maps.EqualFunc(a, b, func(x, y int) bool { return true })) }",
        vec!["true"]
    ),
    maps_equal_func_string_case_insensitive => (
        "package main; import \"fmt\"; import \"maps\"; import \"strings\"; func main() { a := map[int]string{1: \"Go\"}; b := map[int]string{1: \"go\"}; eq := maps.EqualFunc(a, b, func(x, y string) bool { return strings.EqualFold(x, y) }); fmt.Println(eq) }",
        vec!["true"]
    ),
    maps_keys_values_same_map_sizes => (
        "package main; import \"fmt\"; import \"maps\"; func main() { m := map[int]int{1: 10, 2: 20, 3: 30, 4: 40}; fmt.Println(len(maps.Keys(m))); fmt.Println(len(maps.Values(m))) }",
        vec!["4", "4"]
    ),
    maps_keys_after_mutation => (
        "package main; import \"fmt\"; import \"maps\"; func main() { m := map[string]int{\"a\": 1}; m[\"b\"] = 2; fmt.Println(len(maps.Keys(m))) }",
        vec!["2"]
    ),
    maps_values_after_delete => (
        "package main; import \"fmt\"; import \"maps\"; func main() { m := map[int]int{1: 1, 2: 2, 3: 3}; delete(m, 2); fmt.Println(len(maps.Values(m))) }",
        vec!["2"]
    ),
    maps_equal_self_reference_semantics => (
        "package main; import \"fmt\"; import \"maps\"; func main() { m := map[int]int{1: 1}; fmt.Println(maps.Equal(m, m)) }",
        vec!["true"]
    ),
    maps_equal_func_length_mismatch => (
        "package main; import \"fmt\"; import \"maps\"; func main() { a := map[int]int{1: 1, 2: 2}; b := map[int]int{1: 1}; fmt.Println(maps.EqualFunc(a, b, func(x, y int) bool { return true })) }",
        vec!["false"]
    ),
    maps_keys_rune_key_type => (
        "package main; import \"fmt\"; import \"maps\"; func main() { m := map[rune]int{'a': 1, 'b': 2}; fmt.Println(len(maps.Keys(m))) }",
        vec!["2"]
    ),
    maps_values_slice_value_type => (
        "package main; import \"fmt\"; import \"maps\"; func main() { m := map[int][]int{1: {1, 2}, 2: {3}}; total := 0; for v := range maps.Values(m) { total += len(v) }; fmt.Println(total) }",
        vec!["3"]
    ),
    maps_equal_extra_key_in_second => (
        "package main; import \"fmt\"; import \"maps\"; func main() { a := map[string]int{\"only\": 1}; b := map[string]int{\"only\": 1, \"extra\": 2}; fmt.Println(maps.Equal(a, b)) }",
        vec!["false"]
    ),
    maps_equal_func_always_true_same_keys_diff_vals => (
        "package main; import \"fmt\"; import \"maps\"; func main() { a := map[int]int{1: 100}; b := map[int]int{1: 200}; fmt.Println(maps.EqualFunc(a, b, func(x, y int) bool { return true })) }",
        vec!["true"]
    ),
    maps_keys_values_large_map_count => (
        "package main; import \"fmt\"; import \"maps\"; func main() { m := map[int]int{}; for i := 0; i < 10; i++ { m[i] = i * 10 }; fmt.Println(len(maps.Keys(m))); fmt.Println(len(maps.Values(m))) }",
        vec!["10", "10"]
    ),
}

go_compile_cases! {
    maps_keys_float64_key => "package main; import \"maps\"; func main() { _ = maps.Keys(map[float64]int{1.5: 1}) }",
    maps_values_struct_value => "package main; import \"maps\"; type S struct { N int }; func main() { _ = maps.Values(map[int]S{1: {N: 1}}) }",
    maps_equal_pointer_values => "package main; import \"maps\"; func main() { x := 1; a := map[int]*int{1: &x}; b := map[int]*int{1: &x}; _ = maps.Equal(a, b) }",
    maps_equal_func_struct_custom => "package main; import \"maps\"; type P struct { N int }; func main() { a := map[string]P{\"a\": {1}}; b := map[string]P{\"a\": {1}}; _ = maps.EqualFunc(a, b, func(x, y P) bool { return x.N == y.N }) }",
    maps_keys_nested_map_value => "package main; import \"maps\"; func main() { _ = maps.Keys(map[string]map[int]int{\"a\": {1: 1}}) }",
    maps_values_chan_value_type => "package main; import \"maps\"; func main() { _ = maps.Values(map[int]chan int{1: make(chan int, 1)}) }",
    maps_equal_func_nil_vs_empty => "package main; import \"maps\"; func main() { var a map[int]int; _ = maps.EqualFunc(a, map[int]int{}, func(x, y int) bool { return x == y }) }",
    maps_keys_after_clear_reassign => "package main; import \"maps\"; func main() { m := map[int]int{1: 1, 2: 2}; m = map[int]int{}; _ = maps.Keys(m) }",
    maps_values_interface_value => "package main; import \"maps\"; func main() { _ = maps.Values(map[int]interface{}{1: 42, 2: \"x\"}) }",
    maps_equal_int8_keys => "package main; import \"maps\"; func main() { _ = maps.Equal(map[int8]int{1: 1}, map[int8]int{1: 1}) }",
    maps_equal_func_uint_values => "package main; import \"maps\"; func main() { _ = maps.EqualFunc(map[int]uint{1: 5}, map[int]uint{1: 5}, func(x, y uint) bool { return x == y }) }",
    maps_keys_comparable_struct_key => "package main; import \"maps\"; type K struct { A int }; func main() { _ = maps.Keys(map[K]string{{1}: \"v\"}) }",
    maps_values_array_value => "package main; import \"maps\"; func main() { _ = maps.Values(map[int][2]int{1: {1, 2}}) }",
    maps_equal_one_key_missing => "package main; import \"maps\"; func main() { _ = maps.Equal(map[int]int{1: 1, 2: 2}, map[int]int{1: 1, 2: 9}) }",
    maps_equal_func_key_present_value_diff => "package main; import \"maps\"; func main() { _ = maps.EqualFunc(map[string]bool{\"a\": true}, map[string]bool{\"a\": false}, func(x, y bool) bool { return x == y }) }",
    maps_keys_uint_key => "package main; import \"maps\"; func main() { _ = maps.Keys(map[uint]int{1: 1, 2: 2}) }",
}
