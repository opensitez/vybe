use crate::helpers::*;

macro_rules! go_compile_test {
    ($name:ident, $src:expr) => {
        #[test]
        fn $name() {
            compile_ok($src);
        }
    };
}

macro_rules! go_run_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            let out = run_prints($src);
            assert_eq!(out, $expected);
        }
    };
}

macro_rules! run_cases {
    ($( $name:ident => ($src:expr, $expected:expr), )*) => {
        $( go_run_test!($name, $src, $expected); )*
    };
}

macro_rules! compile_cases {
    ($( $name:ident => $src:expr, )*) => {
        $( go_compile_test!($name, $src); )*
    };
}

run_cases! {
    if_with_else_init_runtime => ("package main; import \"fmt\"; func main() { if n := 4; n%2 == 0 { fmt.Println(\"even\") } else { fmt.Println(\"odd\") } }", vec!["even"]),
    switch_with_multiple_values_runtime => ("package main; import \"fmt\"; func main() { x := 3; switch x { case 1, 2: fmt.Println(\"low\"); case 3, 4: fmt.Println(\"mid\") } }", vec!["mid"]),
    for_with_post_assignment_runtime => ("package main; import \"fmt\"; func main() { for i := 1; i < 8; i = i * 2 { fmt.Println(i) } }", vec!["1", "2", "4"]),
    for_with_multiple_init_vars_runtime => ("package main; import \"fmt\"; func main() { for i, j := 0, 3; i < 3; i, j = i+1, j-1 { fmt.Println(i + j) } }", vec!["3", "3", "3"]),
    range_over_array_runtime => ("package main; import \"fmt\"; func main() { values := [3]int{5, 6, 7}; for _, v := range values { fmt.Println(v) } }", vec!["5", "6", "7"]),
    range_over_slice_index_only_runtime => ("package main; import \"fmt\"; func main() { values := []int{8, 9}; for i := range values { fmt.Println(i) } }", vec!["0", "1"]),
    switch_inside_for_runtime => ("package main; import \"fmt\"; func main() { for i := 0; i < 3; i++ { switch i { case 0: fmt.Println(\"zero\"); default: fmt.Println(\"other\") } } }", vec!["zero", "other", "other"]),
    if_inside_switch_runtime => ("package main; import \"fmt\"; func main() { x := 5; switch { case x > 0: if x > 3 { fmt.Println(\"big\") } } }", vec!["big"]),
    for_with_condition_only_runtime => ("package main; import \"fmt\"; func main() { n := 0; for n < 3 { fmt.Println(n); n++ } }", vec!["0", "1", "2"]),
    for_with_omitted_condition_break_runtime => ("package main; import \"fmt\"; func main() { n := 0; for { if n == 2 { break }; fmt.Println(n); n++ } }", vec!["0", "1"]),
    break_from_switch_in_loop_runtime => ("package main; import \"fmt\"; func main() { for i := 0; i < 2; i++ { switch i { case 0: fmt.Println(\"zero\"); break; default: fmt.Println(\"one\") } } }", vec!["zero", "one"]),
    continue_in_inner_loop_runtime => ("package main; import \"fmt\"; func main() { for i := 0; i < 2; i++ { for j := 0; j < 3; j++ { if j == 1 { continue }; fmt.Println(i + j) } } }", vec!["0", "2", "1", "3"]),
    if_with_function_call_init_runtime => ("package main; import \"fmt\"; func value() int { return 7 }; func main() { if n := value(); n > 5 { fmt.Println(n) } }", vec!["7"]),
    switch_with_expressionless_true_runtime => ("package main; import \"fmt\"; func main() { n := 12; switch { case n < 10: fmt.Println(\"small\"); case n < 20: fmt.Println(\"medium\") } }", vec!["medium"]),
    for_with_tuple_post_runtime => ("package main; import \"fmt\"; func main() { for i, j := 0, 2; i < 3; i, j = i+1, j+2 { fmt.Println(i + j) } }", vec!["2", "5", "8"]),
    short_decl_in_for_init_runtime => ("package main; import \"fmt\"; func main() { for n := 1; n <= 3; n++ { fmt.Println(n) } }", vec!["1", "2", "3"]),
    nested_if_with_else_if_and_init_runtime => ("package main; import \"fmt\"; func main() { if n := 5; n < 0 { fmt.Println(\"neg\") } else if n < 10 { if n%2 == 1 { fmt.Println(\"odd\") } } }", vec!["odd"]),
    switch_with_tagless_boolean_cases_runtime => ("package main; import \"fmt\"; func main() { n := 8; switch { case n%3 == 0: fmt.Println(\"three\"); case n%4 == 0: fmt.Println(\"four\") } }", vec!["four"]),
    range_over_nil_slice_runtime => ("package main; import \"fmt\"; func main() { var values []int; for _, v := range values { fmt.Println(v) }; fmt.Println(len(values)) }", vec!["0"]),
    switch_with_default_only_runtime => ("package main; import \"fmt\"; func main() { switch 99 { default: fmt.Println(\"default\") } }", vec!["default"]),
    if_with_negated_bool_runtime => ("package main; import \"fmt\"; func main() { ok := false; if !ok { fmt.Println(\"no\") } }", vec!["no"]),
    if_with_else_branch_runtime => ("package main; import \"fmt\"; func main() { if value := 1; value > 3 { fmt.Println(\"high\") } else { fmt.Println(\"low\") } }", vec!["low"]),
    range_over_string_index_only_runtime => ("package main; import \"fmt\"; func main() { for i := range \"go\" { fmt.Println(i) } }", vec!["0", "1"]),
    nested_switch_runtime => ("package main; import \"fmt\"; func main() { x := 2; switch x { case 2: switch { case x%2 == 0: fmt.Println(\"even-two\") } } }", vec!["even-two"]),
}

compile_cases! {
    switch_with_empty_tag_compile => "package main; func main() { n := 2; switch { case n > 1: _ = n } }",
    switch_with_fallthrough_compile => "package main; func main() { switch 1 { case 1: fallthrough; case 2: } }",
    break_nested_loop_with_label_compile => "package main; func main() { Outer: for i := 0; i < 3; i++ { for j := 0; j < 3; j++ { _ = i + j; break Outer } } }",
    continue_outer_loop_with_label_compile => "package main; func main() { Outer: for i := 0; i < 2; i++ { for j := 0; j < 2; j++ { _ = i + j; continue Outer } } }",
    goto_forward_label_compile => "package main; func main() { goto Done; Done: }",
    goto_backward_label_compile => "package main; func main() { i := 0; Loop: i++; if i < 2 { goto Loop } }",
    labeled_statement_compile => "package main; func main() { Here: ; _ = 1 }",
    select_with_default_compile => "package main; func main() { ch := make(chan int); select { case <-ch: default: } }",
    select_with_receive_case_compile => "package main; func main() { ch := make(chan int); select { case v := <-ch: _ = v; default: } }",
    select_with_send_case_compile => "package main; func main() { ch := make(chan int, 1); select { case ch <- 1: default: } }",
    select_with_assignment_case_compile => "package main; func main() { ch := make(chan int, 1); select { case x, ok := <-ch: _, _ = x, ok; default: } }",
    select_with_short_declaration_case_compile => "package main; func main() { ch := make(chan int, 1); select { case v := <-ch: _ = v; default: } }",
    range_over_map_keys_compile => "package main; func main() { values := map[string]int{\"a\": 1}; for k := range values { _ = k } }",
    range_over_map_values_compile => "package main; func main() { values := map[string]int{\"a\": 1}; for _, v := range values { _ = v } }",
    fallthrough_to_default_compile => "package main; func main() { switch 1 { case 1: fallthrough; default: } }",
    label_before_for_compile => "package main; func main() { Start: for i := 0; i < 1; i++ { _ = i; break Start } }",
    label_before_switch_compile => "package main; func main() { Start: switch 1 { case 1: break Start } }",
    range_over_array_pointer_compile => "package main; func main() { values := &[2]int{1, 2}; for _, v := range values { _ = v } }",
    range_blank_identifiers_compile => "package main; func main() { values := []int{1, 2}; for _, _ = range values { } }",
    select_empty_compile => "package main; func wait() { select {} }; func main() { _ = wait }",
    range_over_string_runes_compile => "package main; func main() { for _, r := range \"hello\" { _ = r } }",
    goto_named_loop_exit_compile => "package main; func main() { i := 0; Loop: if i == 1 { goto Exit }; i++; goto Loop; Exit: }",
    if_with_scoped_short_decl_compile => "package main; func main() { if value := 1; value == 1 { _ = value } }",
    switch_with_scoped_short_decl_compile => "package main; func main() { switch value := 2; value { case 2: _ = value } }",
    for_with_blank_condition_compile => "package main; func main() { for ; true; { break } }",
    labeled_continue_through_nested_switch_compile => "package main; func main() { Outer: for i := 0; i < 1; i++ { switch i { case 0: continue Outer } } }",
}
