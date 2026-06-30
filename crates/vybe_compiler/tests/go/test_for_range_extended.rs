//! Extended `for range` forms: int (Go 1.22), slice index/value, string runes,
//! map keys-only, channel receive, break/continue, and nil-map compile smoke.
//! Distinct from `test_range_over_int.rs`, `test_range_iteration_extra.rs`,
//! and `test_channel_close_range.rs` (close/ok idioms).

use crate::helpers::*;

go_run_cases! {
    range_int_print_last_index_seven => (
        "package main; import \"fmt\"; func main() { last := -1; for i := range 7 { last = i }; fmt.Println(last) }",
        vec!["6"]
    ),
    range_int_skip_first_with_continue => (
        "package main; import \"fmt\"; func main() { total := 0; for i := range 5 { if i == 0 { continue }; total += i }; fmt.Println(total) }",
        vec!["10"]
    ),
    range_int_break_when_sum_exceeds => (
        "package main; import \"fmt\"; func main() { total := 0; for i := range 10 { total += i; if total > 5 { break } }; fmt.Println(total) }",
        vec!["6"]
    ),
    range_int_nested_product_indices => (
        "package main; import \"fmt\"; func main() { product := 1; for i := range 3 { for j := range 2 { if i == 0 && j == 0 { continue }; product *= (i + 1) } }; fmt.Println(product) }",
        vec!["4"]
    ),
    range_int_with_outer_var_shadow => (
        "package main; import \"fmt\"; func main() { bound := 4; total := 0; for i := range bound { total += i }; fmt.Println(total) }",
        vec!["6"]
    ),
    range_slice_index_and_value_pairs => (
        "package main; import \"fmt\"; func main() { sum := 0; for i, v := range []int{10, 20, 30} { sum += i + v }; fmt.Println(sum) }",
        vec!["63"]
    ),
    range_slice_index_only_triple => (
        "package main; import \"fmt\"; func main() { total := 0; for i := range []int{5, 5, 5} { total += i }; fmt.Println(total) }",
        vec!["3"]
    ),
    range_slice_value_only_strings => (
        "package main; import \"fmt\"; func main() { total := 0; for _, word := range []string{\"go\", \"vybe\"} { total += len(word) }; fmt.Println(total) }",
        vec!["6"]
    ),
    range_slice_modify_value_via_index => (
        "package main; import \"fmt\"; func main() { nums := []int{1, 2, 3}; for i := range nums { nums[i] *= 2 }; fmt.Println(nums[1]) }",
        vec!["4"]
    ),
    range_slice_break_on_index_two => (
        "package main; import \"fmt\"; func main() { count := 0; for i := range []int{9, 8, 7, 6} { count++; if i == 2 { break } }; fmt.Println(count) }",
        vec!["3"]
    ),
    range_slice_continue_skip_negative => (
        "package main; import \"fmt\"; func main() { total := 0; for _, v := range []int{1, -1, 2, -2, 3} { if v < 0 { continue }; total += v }; fmt.Println(total) }",
        vec!["6"]
    ),
    range_array_index_value_sum => (
        "package main; import \"fmt\"; func main() { total := 0; for i, v := range [4]int{2, 4, 6, 8} { total += i * v }; fmt.Println(total) }",
        vec!["36"]
    ),
    range_string_rune_sum_ascii => (
        "package main; import \"fmt\"; func main() { total := 0; for _, r := range \"Go\" { total += int(r) }; fmt.Println(total) }",
        vec!["135"]
    ),
    range_string_index_only_len => (
        "package main; import \"fmt\"; func main() { count := 0; for i := range \"hello\" { count = i + 1 }; fmt.Println(count) }",
        vec!["5"]
    ),
    range_string_unicode_rune_count => (
        "package main; import \"fmt\"; func main() { count := 0; for range \"日本\" { count++ }; fmt.Println(count) }",
        vec!["2"]
    ),
    range_string_blank_index_sum_runes => (
        "package main; import \"fmt\"; func main() { total := 0; for _, r := range \"ab\" { total += int(r) }; fmt.Println(total) }",
        vec!["195"]
    ),
    range_map_keys_only_sum => (
        "package main; import \"fmt\"; func main() { total := 0; for k := range map[string]int{\"a\": 1, \"b\": 2, \"c\": 3} { total += len(k) }; fmt.Println(total) }",
        vec!["3"]
    ),
    range_map_keys_only_count => (
        "package main; import \"fmt\"; func main() { count := 0; for range map[int]bool{1: true, 2: false, 3: true} { count++ }; fmt.Println(count) }",
        vec!["3"]
    ),
    range_map_key_only_break_early => (
        "package main; import \"fmt\"; func main() { count := 0; for range map[string]int{\"x\": 1, \"y\": 2} { count++; break }; fmt.Println(count) }",
        vec!["1"]
    ),
    range_map_values_with_blank_key => (
        "package main; import \"fmt\"; func main() { total := 0; for _, v := range map[string]int{\"p\": 4, \"q\": 5} { total += v }; fmt.Println(total) }",
        vec!["9"]
    ),
    range_nil_map_runtime_zero_iters => (
        "package main; import \"fmt\"; func main() { var m map[int]int; count := 0; for range m { count++ }; fmt.Println(count) }",
        vec!["0"]
    ),
    range_empty_slice_zero_iters => (
        "package main; import \"fmt\"; func main() { count := 0; for _, v := range []int{} { count += v }; fmt.Println(count) }",
        vec!["0"]
    ),
    range_string_empty_zero_iters => (
        "package main; import \"fmt\"; func main() { count := 0; for range \"\" { count++ }; fmt.Println(count) }",
        vec!["0"]
    ),
    range_slice_of_bytes_sum => (
        "package main; import \"fmt\"; func main() { total := 0; for _, b := range []byte{'a', 'b', 'c'} { total += int(b) }; fmt.Println(total) }",
        vec!["294"]
    ),
    range_byte_slice_index_value => (
        "package main; import \"fmt\"; func main() { last := byte(0); for i, b := range []byte{10, 20} { if i == 1 { last = b } }; fmt.Println(int(last)) }",
        vec!["20"]
    ),
    range_nested_slice_rows => (
        "package main; import \"fmt\"; func main() { total := 0; for _, row := range [][]int{{1, 2}, {3}} { for _, v := range row { total += v } }; fmt.Println(total) }",
        vec!["6"]
    ),
    range_int_labeled_continue_skip => (
        "package main; import \"fmt\"; func main() { total := 0; outer: for i := range 5 { if i%2 == 0 { continue outer }; total += i }; fmt.Println(total) }",
        vec!["6"]
    ),
    range_slice_labeled_break_outer => (
        "package main; import \"fmt\"; func main() { total := 0; outer: for _, v := range []int{1, 2, 3, 4} { total += v; if v == 2 { break outer } }; fmt.Println(total) }",
        vec!["3"]
    ),
    range_int_count_with_blank => (
        "package main; import \"fmt\"; func main() { count := 0; for range 8 { count++ }; fmt.Println(count) }",
        vec!["8"]
    ),
    range_string_first_rune => (
        "package main; import \"fmt\"; func main() { first := rune(0); for _, r := range \"z\" { first = r; break }; fmt.Println(int(first)) }",
        vec!["122"]
    ),
    range_map_int_keys_only_max => (
        "package main; import \"fmt\"; func main() { max := 0; for k := range map[int]string{3: \"c\", 7: \"g\", 1: \"a\"} { if k > max { max = k } }; fmt.Println(max) }",
        vec!["7"]
    ),
    range_struct_slice_field_index_value => (
        "package main; import \"fmt\"; type bag struct { items []int }; func main() { b := bag{items: []int{4, 5}}; total := 0; for i, v := range b.items { total += i + v }; fmt.Println(total) }",
        vec!["10"]
    ),
    range_int_decrement_pattern => (
        "package main; import \"fmt\"; func main() { acc := 0; for i := range 4 { acc = acc*10 + (3 - i) }; fmt.Println(acc) }",
        vec!["3210"]
    ),
    range_slice_reassign_loop_var => (
        "package main; import \"fmt\"; func main() { sum := 0; for _, v := range []int{1, 2, 3} { v = v * 10; sum += v }; fmt.Println(sum) }",
        vec!["60"]
    ),
    range_channel_buffered_sum => (
        "package main; import \"fmt\"; func main() { ch := make(chan int, 3); ch <- 2; ch <- 4; ch <- 6; close(ch); total := 0; for v := range ch { total += v }; fmt.Println(total) }",
        vec!["12"]
    ),
    range_channel_single_value => (
        "package main; import \"fmt\"; func main() { ch := make(chan int, 1); ch <- 99; close(ch); last := 0; for v := range ch { last = v }; fmt.Println(last) }",
        vec!["99"]
    ),
}

go_compile_cases! {
    range_int_with_defer_in_body_compile =>
        "package main; func main() { for i := range 2 { defer func() { _ = i }() } }",
    range_int_switch_per_iteration_compile =>
        "package main; func main() { for i := range 3 { switch i { case 0: _ = i; case 1: _ = i; default: _ = i } } }",
    range_slice_assign_existing_index_value_compile =>
        "package main; func main() { s := []int{1, 2}; var i int; var v int; for i, v = range s { _, _ = i, v } }",
    range_map_key_only_with_break_compile =>
        "package main; func main() { for k := range map[string]int{\"a\": 1} { _ = k; break } }",
    range_nil_map_key_only_compile =>
        "package main; func main() { var m map[string]int; for k := range m { _ = k } }",
    range_nil_map_key_value_compile =>
        "package main; func main() { var m map[int]string; for k, v := range m { _, _ = k, v } }",
    range_channel_unbuffered_compile =>
        "package main; func main() { ch := make(chan int); go func() { ch <- 1; close(ch) }(); for v := range ch { _ = v } }",
    range_string_bytes_via_conversion_compile =>
        "package main; func main() { for i, b := range []byte(\"go\") { _, _ = i, b } }",
    range_over_rune_slice_compile =>
        "package main; func main() { runes := []rune(\"xy\"); for i, r := range runes { _, _ = i, r } }",
    range_int_nested_labeled_break_compile =>
        "package main; func main() { outer: for i := range 2 { for j := range 2 { if j == 1 { break outer }; _ = i } } }",
    range_slice_nested_continue_compile =>
        "package main; func main() { for _, row := range [][]int{{1, 2}, {3}} { for _, v := range row { if v == 2 { continue }; _ = v } } }",
    range_map_in_if_condition_compile =>
        "package main; func main() { m := map[int]int{1: 2}; for k := range m { if k > 0 { _ = m[k] } } }",
    range_int_in_function_literal_compile =>
        "package main; func main() { fn := func() { for i := range 2 { _ = i } }; fn() }",
    range_over_pointer_to_array_compile =>
        "package main; func main() { arr := &[3]int{1, 2, 3}; for i, v := range arr { _, _ = i, v } }",
    range_over_named_slice_type_compile =>
        "package main; type digits []int; func main() { d := digits{1, 2}; for i, v := range d { _, _ = i, v } }",
    range_string_with_if_break_compile =>
        "package main; func main() { for i, r := range \"abc\" { if i == 1 { break }; _ = r } }",
    range_int_with_short_decl_shadow_compile =>
        "package main; func main() { for i := range 2 { x := i + 1; _ = x } }",
    range_map_range_nested_compile =>
        "package main; func main() { outer := map[string]map[int]bool{\"a\": {1: true}}; for _, inner := range outer { for k := range inner { _ = k } } }",
    range_channel_with_select_compile =>
        "package main; func main() { ch := make(chan int, 1); ch <- 1; close(ch); for v := range ch { select { default: _ = v } } }",
    range_int_continue_labeled_compile =>
        "package main; func main() { loop: for i := range 3 { if i == 1 { continue loop }; _ = i } }",
    range_slice_with_range_int_nested_compile =>
        "package main; func main() { for i := range 2 { for _, v := range []int{i, i + 1} { _ = v } } }",
    range_over_interface_slice_compile =>
        "package main; func main() { items := []interface{}{1, \"x\"}; for i, v := range items { _, _ = i, v } }",
    range_map_string_to_slice_compile =>
        "package main; func main() { m := map[string][]int{\"a\": {1}}; for k, v := range m { _, _ = k, len(v) } }",
    range_int_expression_bound_multiply_compile =>
        "package main; func main() { for i := range 2 * 2 { _ = i } }",
    range_string_utf8_multibyte_compile =>
        "package main; func main() { for i, r := range \"é\" { _, _ = i, r } }",
    range_slice_header_after_append_compile =>
        "package main; func main() { s := []int{1}; s = append(s, 2); for i, v := range s { _, _ = i, v } }",
    range_int_zero_bound_compile =>
        "package main; func main() { for i := range 0 { _ = i } }",
    range_map_delete_during_iteration_compile =>
        "package main; func main() { m := map[string]int{\"a\": 1, \"b\": 2}; for k := range m { delete(m, k); break } }",
    range_channel_closed_empty_compile =>
        "package main; func main() { ch := make(chan struct{}); close(ch); for range ch { } }",
}
