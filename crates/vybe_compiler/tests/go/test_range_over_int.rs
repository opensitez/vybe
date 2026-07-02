//! Go 1.22+ integer range (`for i := range n`) — indices, break/continue, sums.
//! Distinct from `test_range_iteration_extra.rs` (slice/array/map/string/channel range).


go_run_cases! {
    range_int_print_indices => (
        "package main; import \"fmt\"; func main() { for i := range 4 { fmt.Println(i) } }",
        vec!["0", "1", "2", "3"]
    ),
    range_int_sum_five => (
        "package main; import \"fmt\"; func main() { total := 0; for i := range 5 { total += i }; fmt.Println(total) }",
        vec!["10"]
    ),
    range_int_sum_ten => (
        "package main; import \"fmt\"; func main() { total := 0; for i := range 10 { total += i }; fmt.Println(total) }",
        vec!["45"]
    ),
    range_int_zero_iterations => (
        "package main; import \"fmt\"; func main() { count := 0; for i := range 0 { count += i + 1 }; fmt.Println(count) }",
        vec!["0"]
    ),
    range_int_negative_no_iterations => (
        "package main; import \"fmt\"; func main() { count := 0; for i := range -3 { count += i + 1 }; fmt.Println(count) }",
        vec!["0"]
    ),
    range_int_single_iteration => (
        "package main; import \"fmt\"; func main() { last := -1; for i := range 1 { last = i }; fmt.Println(last) }",
        vec!["0"]
    ),
    range_int_variable_bound => (
        "package main; import \"fmt\"; func main() { n := 6; total := 0; for i := range n { total += i }; fmt.Println(total) }",
        vec!["15"]
    ),
    range_int_expression_bound => (
        "package main; import \"fmt\"; func main() { total := 0; for i := range 2 + 3 { total += i }; fmt.Println(total) }",
        vec!["10"]
    ),
    range_int_break_at_three => (
        "package main; import \"fmt\"; func main() { total := 0; for i := range 8 { if i == 3 { break }; total += i }; fmt.Println(total) }",
        vec!["3"]
    ),
    range_int_break_on_first => (
        "package main; import \"fmt\"; func main() { total := 0; for i := range 5 { total += i; break }; fmt.Println(total) }",
        vec!["0"]
    ),
    range_int_continue_skip_even => (
        "package main; import \"fmt\"; func main() { total := 0; for i := range 5 { if i%2 == 0 { continue }; total += i }; fmt.Println(total) }",
        vec!["4"]
    ),
    range_int_continue_skip_middle => (
        "package main; import \"fmt\"; func main() { total := 0; for i := range 4 { if i == 2 { continue }; total += i }; fmt.Println(total) }",
        vec!["3"]
    ),
    range_int_product_nonzero_indices => (
        "package main; import \"fmt\"; func main() { total := 1; for i := range 4 { if i > 0 { total *= i } }; fmt.Println(total) }",
        vec!["6"]
    ),
    range_int_nested_iteration_count => (
        "package main; import \"fmt\"; func main() { count := 0; for i := range 3 { for j := range 2 { count += i + j } }; fmt.Println(count) }",
        vec!["9"]
    ),
    range_int_conditional_sum_from_two => (
        "package main; import \"fmt\"; func main() { total := 0; for i := range 5 { if i >= 2 { total += i } }; fmt.Println(total) }",
        vec!["9"]
    ),
    range_int_last_index => (
        "package main; import \"fmt\"; func main() { last := -1; for i := range 5 { last = i }; fmt.Println(last) }",
        vec!["4"]
    ),
    range_int_count_with_blank_range => (
        "package main; import \"fmt\"; func main() { count := 0; for range 6 { count++ }; fmt.Println(count) }",
        vec!["6"]
    ),
    range_int_sum_squares => (
        "package main; import \"fmt\"; func main() { total := 0; for i := range 4 { total += i * i }; fmt.Println(total) }",
        vec!["14"]
    ),
    range_int_labeled_break_outer => (
        "package main; import \"fmt\"; func main() { total := 0; outer: for i := range 4 { for j := range 3 { if i == 2 && j == 1 { break outer }; total += 1 } }; fmt.Println(total) }",
        vec!["7"]
    ),
    range_int_function_return_bound => (
        "package main; import \"fmt\"; func bound() int { return 4 }; func main() { total := 0; for i := range bound() { total += i }; fmt.Println(total) }",
        vec!["6"]
    ),
}

go_compile_cases! {
    range_int_with_defer_compile => "package main; func main() { for i := range 3 { defer func() { _ = i }() } }",
    range_int_switch_inside_compile => "package main; func main() { for i := range 4 { switch i { case 0: _ = i; default: _ = i } } }",
    range_int_labeled_continue_outer_compile => "package main; func main() { outer: for i := range 3 { if i == 1 { continue outer }; _ = i } }",
    range_int_nested_break_inner_compile => "package main; func main() { for i := range 2 { for j := range 2 { if j == 1 { break }; _ = j } } }",
    range_int_short_decl_in_body_compile => "package main; func main() { for i := range 2 { next := i + 1; _ = next } }",
    range_int_if_chain_compile => "package main; func main() { for i := range 3 { if i < 1 { _ = i } else if i < 2 { _ = i } else { _ = i } } }",
}
