use crate::helpers::*;

go_run_cases! {
    defer_lifo_four_level_stack =>
        ("package main; import \"fmt\"; func main() { defer fmt.Println(\"a\"); defer fmt.Println(\"b\"); defer fmt.Println(\"c\"); defer fmt.Println(\"d\"); }", vec!["d", "c", "b", "a"]),
    defer_inner_frame_finishes_before_outer =>
        ("package main; import \"fmt\"; func inner() { defer fmt.Println(\"inner\"); }; func main() { defer fmt.Println(\"outer\"); inner(); }", vec!["inner", "outer"]),
    defer_registered_inside_deferred_func_runs_first =>
        ("package main; import \"fmt\"; func main() { defer func() { defer fmt.Println(\"inner\"); fmt.Println(\"outer\") }(); }", vec!["outer", "inner"]),
    defer_lifo_interleaved_with_work_before_panic =>
        ("package main; import \"fmt\"; func run() { defer fmt.Println(\"third\"); defer fmt.Println(\"second\"); defer func() { recover() }(); fmt.Println(\"first\"); panic(\"stop\") }; func main() { run() }", vec!["first", "third", "second"]),
    defer_lifo_mixed_named_funcs_and_literals =>
        ("package main; import \"fmt\"; func mark(label string) { fmt.Println(label) }; func main() { defer mark(\"alpha\"); defer func() { fmt.Println(\"beta\") }(); defer mark(\"gamma\"); }", vec!["gamma", "beta", "alpha"]),
    named_return_string_overwritten_by_defer =>
        ("package main; import \"fmt\"; func greet() (msg string) { defer func() { msg = \"bye\" }(); return \"hi\" }; func main() { fmt.Println(greet()); }", vec!["bye"]),
    named_return_pair_both_mutated_by_defers =>
        ("package main; import \"fmt\"; func stats() (total int, count int) { defer func() { count = 4 }(); defer func() { total = 9 }(); return 1, 2 }; func main() { t, c := stats(); fmt.Println(t); fmt.Println(c); }", vec!["9", "4"]),
    named_return_bool_flipped_by_defer =>
        ("package main; import \"fmt\"; func check() (ok bool) { defer func() { ok = true }(); return false }; func main() { fmt.Println(check()); }", vec!["true"]),
    named_return_explicit_value_replaced_by_defer =>
        ("package main; import \"fmt\"; func build() (n int) { defer func() { n = 99 }(); return 5 }; func main() { fmt.Println(build()); }", vec!["99"]),
    named_return_bare_return_scaled_by_defer =>
        ("package main; import \"fmt\"; func scale() (n int) { defer func() { n = n * 3 }(); n = 4; return }; func main() { fmt.Println(scale()); }", vec!["12"]),
    named_return_sum_incremented_once_by_defer =>
        ("package main; import \"fmt\"; func total() (sum int) { defer func() { sum = sum + 10 }(); return 7 }; func main() { fmt.Println(total()); }", vec!["17"]),
    recover_captures_bool_false_panic =>
        ("package main; import \"fmt\"; func run() { defer func() { value := recover(); fmt.Println(value == false) }(); panic(false) }; func main() { run() }", vec!["true"]),
    recover_captures_float64_panic =>
        ("package main; import \"fmt\"; func run() { defer func() { fmt.Println(recover()) }(); panic(2.5) }; func main() { run() }", vec!["2.5"]),
    recover_captures_struct_field_via_type_assertion =>
        ("package main; import \"fmt\"; type stop struct { code int }; func run() { defer func() { value := recover(); if err, ok := value.(stop); ok { fmt.Println(err.code) } }(); panic(stop{code: 42}) }; func main() { run() }", vec!["42"]),
    recover_nil_panic_yields_nil =>
        ("package main; import \"fmt\"; func run() { defer func() { fmt.Println(recover() == nil) }(); panic(nil) }; func main() { run() }", vec!["true"]),
    recover_uint_panic_value =>
        ("package main; import \"fmt\"; func run() { defer func() { fmt.Println(recover()) }(); panic(uint(8)) }; func main() { run() }", vec!["8"]),
    recover_rune_panic_value =>
        ("package main; import \"fmt\"; func run() { defer func() { fmt.Println(recover()) }(); panic(rune(65)) }; func main() { run() }", vec!["65"]),
    defer_in_range_loop_prints_lifo =>
        ("package main; import \"fmt\"; func main() { for _, value := range []int{10, 20, 30} { defer fmt.Println(value) } }", vec!["30", "20", "10"]),
    defer_in_nested_loops_registers_six_callbacks =>
        ("package main; import \"fmt\"; func main() { for i := 0; i < 2; i++ { for j := 0; j < 3; j++ { defer fmt.Println(i*10 + j) } } }", vec!["12", "11", "10", "2", "1", "0"]),
    defer_in_loop_with_break_registers_three =>
        ("package main; import \"fmt\"; func main() { for i := 0; i < 5; i++ { defer fmt.Println(i); if i == 2 { break } } }", vec!["2", "1", "0"]),
    defer_in_loop_accumulates_counter_on_exit =>
        ("package main; import \"fmt\"; func main() { total := 0; for i := 1; i <= 3; i++ { defer func() { total = total + i }() }; fmt.Println(total) }", vec!["6"]),
    recover_from_plain_function_is_nil =>
        ("package main; import \"fmt\"; func probe() { fmt.Println(recover() == nil) }; func main() { probe() }", vec!["true"]),
    defer_recover_prints_panic_message =>
        ("package main; import \"fmt\"; func run() { defer func() { if r := recover(); r != nil { fmt.Println(r) } }(); panic(\"halt\") }; func main() { run() }", vec!["halt"]),
    two_deferred_recovers_only_first_gets_panic =>
        ("package main; import \"fmt\"; func run() { defer func() { fmt.Println(recover() == nil) }(); defer func() { fmt.Println(recover() != nil) }(); panic(\"boom\") }; func main() { run() }", vec!["true", "true"]),
    panic_in_helper_recovered_by_deferred_closure =>
        ("package main; import \"fmt\"; func boom() { panic(\"fail\") }; func run() { defer func() { if recover() != nil { fmt.Println(\"saved\") } }(); boom() }; func main() { run() }", vec!["saved"]),
    defer_recover_allows_post_panic_code_in_caller =>
        ("package main; import \"fmt\"; func run() { defer func() { recover() }(); panic(\"x\"); fmt.Println(\"skip\") }; func main() { run(); fmt.Println(\"after\") }", vec!["after"]),
    panic_in_inner_defer_caught_by_outer_recover =>
        ("package main; import \"fmt\"; func run() { defer func() { if recover() != nil { fmt.Println(\"caught\") } }(); defer func() { panic(\"inner\") }() }; func main() { run() }", vec!["caught"]),
    defer_recover_with_message_parameter =>
        ("package main; import \"fmt\"; func run() { defer func(label string) { if recover() != nil { fmt.Println(label) } }(\"handled\"); panic(\"err\") }; func main() { run() }", vec!["handled"]),
    defer_lifo_runs_cleanup_before_recover_on_panic =>
        ("package main; import \"fmt\"; func run() { defer fmt.Println(\"cleanup\"); defer func() { recover() }(); panic(\"stop\") }; func main() { run(); fmt.Println(\"done\") }", vec!["cleanup", "done"]),
    recover_type_switch_on_int_panic =>
        ("package main; import \"fmt\"; func run() { defer func() { switch value := recover().(type) { case int: fmt.Println(value + 1); default: fmt.Println(0) } }(); panic(6) }; func main() { run() }", vec!["7"]),
}
