//! Advanced `select` patterns: default branches, multi-case readiness, nil and closed
//! channels, and send-side cases — distinct from `test_channel_select_patterns_extra.rs`.

go_run_cases! {
    select_default_when_nil_receive_blocked =>
        ("package main; import \"fmt\"; func main() { var ch chan int; select { case <-ch: fmt.Println(\"recv\"); default: fmt.Println(\"default\") } }", vec!["default"]),

    select_default_when_nil_send_blocked =>
        ("package main; import \"fmt\"; func main() { var ch chan int; select { case ch <- 1: fmt.Println(\"send\"); default: fmt.Println(\"default\") } }", vec!["default"]),

    select_default_when_unbuffered_receive_blocked =>
        ("package main; import \"fmt\"; func main() { ch := make(chan int); select { case <-ch: fmt.Println(\"recv\"); default: fmt.Println(\"default\") } }", vec!["default"]),

    select_default_with_three_blocked_nil_cases =>
        ("package main; import \"fmt\"; func main() { var a, b, c chan int; select { case <-a: fmt.Println(1); case <-b: fmt.Println(2); case <-c: fmt.Println(3); default: fmt.Println(\"idle\") } }", vec!["idle"]),

    select_closed_channel_receive_zero_not_default =>
        ("package main; import \"fmt\"; func main() { ch := make(chan int); close(ch); select { case v := <-ch: fmt.Println(v); default: fmt.Println(\"default\") } }", vec!["0"]),

    select_closed_channel_receive_ok_false =>
        ("package main; import \"fmt\"; func main() { ch := make(chan int); close(ch); select { case v, ok := <-ch: fmt.Println(v); fmt.Println(ok); default: fmt.Println(\"default\") } }", vec!["0", "false"]),

    select_closed_buffered_drains_value_then_zero =>
        ("package main; import \"fmt\"; func main() { ch := make(chan int, 1); ch <- 42; close(ch); select { case v := <-ch: fmt.Println(v) }; select { case v, ok := <-ch: fmt.Println(v); fmt.Println(ok) } }", vec!["42", "0", "false"]),

    select_receive_wins_over_default_when_buffered_ready =>
        ("package main; import \"fmt\"; func main() { ch := make(chan int, 1); ch <- 11; select { case v := <-ch: fmt.Println(v); default: fmt.Println(0) } }", vec!["11"]),

    select_first_ready_among_three_receive_cases =>
        ("package main; import \"fmt\"; func main() { ch1 := make(chan int, 1); ch2 := make(chan int); ch3 := make(chan int); ch1 <- 3; select { case v := <-ch1: fmt.Println(v); case <-ch2: fmt.Println(2); case <-ch3: fmt.Println(1); default: fmt.Println(0) } }", vec!["3"]),

    select_mixed_nil_and_ready_buffered_receive =>
        ("package main; import \"fmt\"; func main() { var blocked chan int; ready := make(chan int, 1); ready <- 6; select { case <-blocked: fmt.Println(0); case v := <-ready: fmt.Println(v); default: fmt.Println(\"default\") } }", vec!["6"]),

    select_receive_ok_true_on_buffered_value =>
        ("package main; import \"fmt\"; func main() { ch := make(chan int, 1); ch <- 99; select { case v, ok := <-ch: fmt.Println(v); fmt.Println(ok); default: fmt.Println(\"default\") } }", vec!["99", "true"]),

    select_string_channel_receive_ready =>
        ("package main; import \"fmt\"; func main() { ch := make(chan string, 1); ch <- \"go\"; select { case s := <-ch: fmt.Println(s); default: fmt.Println(\"default\") } }", vec!["go"]),

    select_bool_channel_false_value =>
        ("package main; import \"fmt\"; func main() { ch := make(chan bool, 1); ch <- false; select { case b := <-ch: fmt.Println(b); default: fmt.Println(true) } }", vec!["false"]),

    select_receive_discards_value_with_blank =>
        ("package main; import \"fmt\"; func main() { ch := make(chan int, 1); ch <- 17; select { case <-ch: fmt.Println(\"got\"); default: fmt.Println(\"miss\") } }", vec!["got"]),

    select_default_only_no_cases =>
        ("package main; import \"fmt\"; func main() { select { default: fmt.Println(\"only\") } }", vec!["only"]),

    select_nested_default_inner_then_outer =>
        ("package main; import \"fmt\"; func main() { select { default: select { default: fmt.Println(\"inner\") } } }", vec!["inner"]),

    select_two_ready_receives_first_case_wins =>
        ("package main; import \"fmt\"; func main() { a := make(chan int, 1); b := make(chan int, 1); a <- 1; b <- 2; select { case v := <-a: fmt.Println(v); case v := <-b: fmt.Println(v); default: fmt.Println(0) } }", vec!["1"]),

    select_closed_after_draining_buffered_value =>
        ("package main; import \"fmt\"; func main() { ch := make(chan int, 2); ch <- 10; ch <- 20; close(ch); select { case v := <-ch: fmt.Println(v) }; select { case v := <-ch: fmt.Println(v) }; select { case v, ok := <-ch: fmt.Println(v); fmt.Println(ok) } }", vec!["10", "20", "0", "false"]),

    select_nil_plus_closed_channel_prefers_closed =>
        ("package main; import \"fmt\"; func main() { chE := make(chan int); close(chE); var chN chan int; select { case <-chN: fmt.Println(\"nil\"); case v := <-chE: fmt.Println(v); default: fmt.Println(\"default\") } }", vec!["0"]),

    select_default_between_two_blocked_receives =>
        ("package main; import \"fmt\"; func main() { left := make(chan int); right := make(chan int); select { case <-left: fmt.Println(\"left\"); case <-right: fmt.Println(\"right\"); default: fmt.Println(\"neither\") } }", vec!["neither"]),
}

go_compile_cases! {
    select_default_when_unbuffered_send_blocked =>
        "package main; func main() { ch := make(chan int); select { case ch <- 9: default: } }",

    select_default_when_buffered_send_full =>
        "package main; func main() { ch := make(chan int, 1); ch <- 1; select { case ch <- 2: default: } }",

    select_send_wins_over_default_when_buffer_has_space =>
        "package main; func main() { ch := make(chan int, 2); select { case ch <- 7: default: } }",

    select_second_case_send_when_first_receive_blocked =>
        "package main; func main() { recv := make(chan int); send := make(chan int, 1); select { case <-recv: case send <- 5: default: } }",

    select_send_case_to_full_then_default_in_loop =>
        "package main; func main() { ch := make(chan int, 1); ch <- 1; for i := 0; i < 3; i++ { select { case ch <- i: default: } } }",

    select_string_channel_send_then_receive =>
        "package main; func main() { ch := make(chan string, 1); select { case ch <- \"go\": default: }; select { case <-ch: default: } }",

    select_struct_pointer_channel_send_receive =>
        "package main; type node struct { n int }; func main() { ch := make(chan *node, 1); select { case ch <- &node{n: 8}: default: }; select { case <-ch: default: } }",

    select_two_sends_one_full_one_open =>
        "package main; func main() { full := make(chan int, 1); full <- 1; open := make(chan int, 1); select { case full <- 2: case open <- 3: default: } }",

    select_receive_after_prior_send_same_channel =>
        "package main; func main() { ch := make(chan int, 1); select { case ch <- 4: default: }; select { case <-ch: default: } }",

    select_four_mixed_cases_with_default =>
        "package main; func main() { var nilCh chan int; ready := make(chan int, 1); send := make(chan int, 1); recv := make(chan int); select { case <-nilCh: case v := <-ready: _ = v; case send <- 1: case <-recv: default: } }",

    select_closed_and_nil_receive_cases =>
        "package main; func main() { closed := make(chan int); close(closed); var nilCh chan int; select { case <-closed: case <-nilCh: default: } }",

    select_send_only_typed_channel_in_case =>
        "package main; func main() { ch := make(chan int, 1); var out chan<- int = ch; select { case out <- 2: default: } }",

    select_recv_only_typed_channel_in_case =>
        "package main; func main() { ch := make(chan int, 1); ch <- 1; var in <-chan int = ch; select { case <-in: default: } }",

    select_assign_to_outer_var_in_case =>
        "package main; func main() { ch := make(chan int, 1); ch <- 5; var result int; select { case v := <-ch: result = v; default: }; _ = result }",

    select_in_helper_function_returning_tag =>
        "package main; func pick(ch chan int) int { select { case v := <-ch: return v; default: return -1 } }; func main() { ch := make(chan int, 1); ch <- 2; _ = pick(ch) }",

    select_with_labeled_break_from_loop =>
        "package main; func main() { ch := make(chan int, 1); loop: for { select { case <-ch: break loop; default: return } } }",

    select_multiple_nil_sends_blocked =>
        "package main; func main() { var a, b chan int; select { case a <- 1: case b <- 2: default: } }",
}
