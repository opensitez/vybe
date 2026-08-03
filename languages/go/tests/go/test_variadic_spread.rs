//! Variadic parameters, slice spread (`...`), forwarding, and mixed fixed/variadic calls.
//!
//! Distinct from `test_variadic.rs` (basic sum/count/join) and
//! `test_functions_patterns_extra.rs` (single forward/spread smoke tests).

go_run_cases! {
    spread_empty_int_slice_variadic_zero_sum => (
        "package main; import \"fmt\"; func sum(nums ...int) int { total := 0; for _, n := range nums { total += n }; return total }; func main() { fmt.Println(sum([]int{}...)); }",
        vec!["0"]
    ),
    spread_single_element_int_slice => (
        "package main; import \"fmt\"; func sum(nums ...int) int { total := 0; for _, n := range nums { total += n }; return total }; func main() { fmt.Println(sum([]int{9}...)); }",
        vec!["9"]
    ),
    spread_int_slice_after_fixed_multiplier => (
        "package main; import \"fmt\"; func scale(factor int, nums ...int) int { total := 0; for _, n := range nums { total += n * factor }; return total }; func main() { batch := []int{2, 3, 4}; fmt.Println(scale(10, batch...)); }",
        vec!["90"]
    ),
    spread_string_slice_after_fixed_prefix => (
        "package main; import \"fmt\"; func tag(prefix string, words ...string) { for _, w := range words { fmt.Println(prefix + w) } }; func main() { rest := []string{\"go\", \"vybe\"}; tag(\">\", rest...); }",
        vec![">go", ">vybe"]
    ),
    spread_slice_returned_from_helper => (
        "package main; import \"fmt\"; func digits() []int { return []int{1, 2, 3} }; func sum(nums ...int) int { total := 0; for _, n := range nums { total += n }; return total }; func main() { fmt.Println(sum(digits()...)); }",
        vec!["6"]
    ),
    spread_subslice_into_variadic => (
        "package main; import \"fmt\"; func sum(nums ...int) int { total := 0; for _, n := range nums { total += n }; return total }; func main() { all := []int{10, 1, 2, 3}; fmt.Println(sum(all[1:]...)); }",
        vec!["6"]
    ),
    forward_variadic_single_delegate => (
        "package main; import \"fmt\"; func sink(nums ...int) int { total := 0; for _, n := range nums { total += n }; return total }; func relay(nums ...int) int { return sink(nums...) }; func main() { fmt.Println(relay(4, 5)); }",
        vec!["9"]
    ),
    forward_variadic_double_delegate_chain => (
        "package main; import \"fmt\"; func end(nums ...int) int { return len(nums) }; func mid(nums ...int) int { return end(nums...) }; func start(nums ...int) int { return mid(nums...) }; func main() { fmt.Println(start(1, 2, 3, 4)); }",
        vec!["4"]
    ),
    forward_variadic_injects_count_before_delegate => (
        "package main; import \"fmt\"; func emit(nums ...int) { for _, n := range nums { fmt.Println(n) } }; func relay(nums ...int) { fmt.Println(len(nums)); emit(nums...) }; func main() { relay(7, 8); }",
        vec!["2", "7", "8"]
    ),
    forward_variadic_to_fmt_println_wrapper => (
        "package main; import \"fmt\"; func show(parts ...interface{}) { fmt.Println(parts...) }; func main() { show(\"vybe\", 42); }",
        vec!["vybe 42"]
    ),
    mixed_two_fixed_plus_variadic_offset_sum => (
        "package main; import \"fmt\"; func tally(base int, step int, nums ...int) int { total := base; for _, n := range nums { total += n + step }; return total }; func main() { fmt.Println(tally(100, 1, 2, 3)); }",
        vec!["108"]
    ),
    mixed_three_fixed_strings_bracket_variadic => (
        "package main; import \"fmt\"; func bracket(open string, close string, sep string, parts ...string) string { out := open; for i, p := range parts { if i > 0 { out += sep }; out += p }; return out + close }; func main() { fmt.Println(bracket(\"[\", \"]\", \",\", \"a\", \"b\")); }",
        vec!["[a,b]"]
    ),
    mixed_literals_plus_spread_string_slice => (
        "package main; import \"fmt\"; func join3(a string, b string, rest ...string) int { return len(rest) + len(a) + len(b) }; func main() { tail := []string{\"c\", \"d\"}; fmt.Println(join3(\"x\", \"y\", tail...)); }",
        vec!["4"]
    ),
    variadic_int_product => (
        "package main; import \"fmt\"; func product(nums ...int) int { p := 1; for _, n := range nums { p *= n }; return p }; func main() { fmt.Println(product(2, 3, 4)); }",
        vec!["24"]
    ),
    variadic_int_minimum => (
        "package main; import \"fmt\"; func minimum(nums ...int) int { m := nums[0]; for _, n := range nums { if n < m { m = n } }; return m }; func main() { fmt.Println(minimum(5, 1, 8, 2)); }",
        vec!["1"]
    ),
    variadic_float64_sum => (
        "package main; import \"fmt\"; func sum(nums ...float64) float64 { total := 0.0; for _, n := range nums { total += n }; return total }; func main() { fmt.Println(sum(0.5, 1.5, 2.0)); }",
        vec!["4"]
    ),
    variadic_empty_strings_count_three => (
        "package main; import \"fmt\"; func count(words ...string) int { return len(words) }; func main() { fmt.Println(count(\"\", \"\", \"\")); }",
        vec!["3"]
    ),
    variadic_interface_mixed_len => (
        "package main; import \"fmt\"; func pack(values ...interface{}) int { return len(values) }; func main() { fmt.Println(pack(1, \"two\", true)); }",
        vec!["3"]
    ) }

go_compile_cases! {
    spread_nil_int_slice_to_variadic => "package main; import \"fmt\"; func sum(nums ...int) int { return len(nums) }; func main() { var s []int; fmt.Println(sum(s...)) }",
    forward_variadic_nested_helper_compile => "package main; func sink(nums ...int) int { return len(nums) }; func relay(nums ...int) int { return func(v ...int) int { return sink(v...) }(nums...) }; func main() { _ = relay(1, 2) }",
    forward_variadic_closure_compile => "package main; func build() func(...int) int { return func(nums ...int) int { return len(nums) } }; func main() { fn := build(); _ = fn(1, 2, 3) }",
    mixed_params_int_string_variadic_compile => "package main; func log(level int, tag string, msgs ...string) int { return level + len(tag) + len(msgs) }; func main() { _ = log(2, \"app\", \"a\", \"b\") }",
    variadic_method_receiver_compile => "package main; type Logger struct{}; func (l Logger) Write(parts ...string) int { return len(parts) }; func main() { _ = Logger{}.Write(\"a\", \"b\") }",
    spread_string_slice_into_variadic_compile => "package main; func take(words ...string) int { return len(words) }; func main() { tail := []string{\"x\", \"y\"}; _ = take(tail...) }",
    forward_variadic_to_peer_with_slice_copy_compile => "package main; func sink(nums ...int) int { return len(nums) }; func relay(nums ...int) int { copy := append([]int(nil), nums...); return sink(copy...) }; func main() { _ = relay(3, 4, 5) }" }
