//! Advanced variadic calls: empty/mixed args, slice spread, interface{} packs,
//! forwarding chains, and compile-only type/signature forms.
//! Distinct from `test_variadic_spread.rs` and `test_functions_patterns_extra.rs`.

go_run_cases! {
    variadic_empty_call_zero_len => (
        "package main; import \"fmt\"; func count(nums ...int) int { return len(nums) }; func main() { fmt.Println(count()) }",
        vec!["0"]
    ),
    variadic_empty_call_string_join => (
        "package main; import \"fmt\"; func join(sep string, parts ...string) string { out := \"\"; for i, p := range parts { if i > 0 { out += sep }; out += p }; return out }; func main() { fmt.Println(join(\",\", )) }",
        vec![""]
    ),
    spread_nil_slice_len_zero => (
        "package main; import \"fmt\"; func size(items ...int) int { return len(items) }; func main() { var s []int; fmt.Println(size(s...)) }",
        vec!["0"]
    ),
    spread_full_slice_three_elements => (
        "package main; import \"fmt\"; func sum(nums ...int) int { t := 0; for _, n := range nums { t += n }; return t }; func main() { batch := []int{1, 2, 3}; fmt.Println(sum(batch...)) }",
        vec!["6"]
    ),
    spread_subslice_middle_two => (
        "package main; import \"fmt\"; func sum(nums ...int) int { t := 0; for _, n := range nums { t += n }; return t }; func main() { all := []int{5, 1, 2, 3, 9}; fmt.Println(sum(all[1:3]...)) }",
        vec!["3"]
    ),
    mixed_one_fixed_plus_variadic => (
        "package main; import \"fmt\"; func prefix(tag string, msgs ...string) int { return len(tag) + len(msgs) }; func main() { fmt.Println(prefix(\"ERR\", \"a\", \"b\")) }",
        vec!["5"]
    ),
    mixed_two_fixed_int_variadic_sum => (
        "package main; import \"fmt\"; func offset(base int, step int, vals ...int) int { t := base; for _, v := range vals { t += v + step }; return t }; func main() { fmt.Println(offset(10, 1, 2, 3)) }",
        vec!["18"]
    ),
    mixed_three_fixed_before_spread => (
        "package main; import \"fmt\"; func frame(a string, b string, c string, rest ...string) int { return len(a) + len(b) + len(c) + len(rest) }; func main() { tail := []string{\"d\"}; fmt.Println(frame(\"x\", \"y\", \"z\", tail...)) }",
        vec!["4"]
    ),
    forward_variadic_to_peer_direct => (
        "package main; import \"fmt\"; func sink(nums ...int) int { t := 0; for _, n := range nums { t += n }; return t }; func relay(nums ...int) int { return sink(nums...) }; func main() { fmt.Println(relay(1, 2, 3)) }",
        vec!["6"]
    ),
    forward_variadic_triple_hop => (
        "package main; import \"fmt\"; func end(words ...string) int { return len(words) }; func mid(words ...string) int { return end(words...) }; func start(words ...string) int { return mid(words...) }; func main() { fmt.Println(start(\"a\", \"b\")) }",
        vec!["2"]
    ),
    forward_variadic_prepends_literal => (
        "package main; import \"fmt\"; func emit(nums ...int) { for _, n := range nums { fmt.Println(n) } }; func relay(nums ...int) { emit(append([]int{0}, nums...)...) }; func main() { relay(5, 6) }",
        vec!["0", "5", "6"]
    ),
    variadic_interface_len_mixed_types => (
        "package main; import \"fmt\"; func pack(values ...interface{}) int { return len(values) }; func main() { fmt.Println(pack(1, \"two\", true, 4.0)) }",
        vec!["4"]
    ),
    variadic_interface_first_element => (
        "package main; import \"fmt\"; func first(values ...interface{}) interface{} { return values[0] }; func main() { fmt.Println(first(99, \"x\")) }",
        vec!["99"]
    ),
    variadic_string_max_length => (
        "package main; import \"fmt\"; func longest(words ...string) int { m := 0; for _, w := range words { if len(w) > m { m = len(w) } }; return m }; func main() { fmt.Println(longest(\"go\", \"vybe\", \"a\")) }",
        vec!["4"]
    ),
    variadic_int_only_last => (
        "package main; import \"fmt\"; func last(nums ...int) int { return nums[len(nums)-1] }; func main() { fmt.Println(last(3, 7, 11)) }",
        vec!["11"]
    ),
    variadic_bool_all_true => (
        "package main; import \"fmt\"; func allTrue(flags ...bool) bool { for _, f := range flags { if !f { return false } }; return true }; func main() { fmt.Println(allTrue(true, true, true)) }",
        vec!["true"]
    ),
    variadic_float64_average_two => (
        "package main; import \"fmt\"; func avg(nums ...float64) float64 { if len(nums) == 0 { return 0 }; s := 0.0; for _, n := range nums { s += n }; return s / float64(len(nums)) }; func main() { fmt.Println(avg(2.0, 4.0)) }",
        vec!["3"]
    ),
    spread_after_literals_in_call => (
        "package main; import \"fmt\"; func sum(nums ...int) int { t := 0; for _, n := range nums { t += n }; return t }; func main() { extra := []int{4, 5}; fmt.Println(sum(1, 2, 3, extra...)) }",
        vec!["15"]
    ),
    spread_string_runes_as_string_variadic => (
        "package main; import \"fmt\"; func concat(parts ...string) string { out := \"\"; for _, p := range parts { out += p }; return out }; func main() { letters := []string{\"a\", \"b\"}; fmt.Println(concat(letters...)) }",
        vec!["ab"]
    ),
    variadic_with_return_count_and_sum => (
        "package main; import \"fmt\"; func stats(nums ...int) (int, int) { t := 0; for _, n := range nums { t += n }; return len(nums), t }; func main() { c, s := stats(2, 3, 4); fmt.Println(c); fmt.Println(s) }",
        vec!["3", "9"]
    ),
    variadic_recursive_count_via_forward => (
        "package main; import \"fmt\"; func depth(level int, tags ...string) int { if level == 0 { return len(tags) }; return depth(level-1, tags...) }; func main() { fmt.Println(depth(2, \"a\", \"b\", \"c\")) }",
        vec!["3"]
    ),
    variadic_method_on_struct => (
        "package main; import \"fmt\"; type Tally struct{}; func (t Tally) Add(nums ...int) int { s := 0; for _, n := range nums { s += n }; return s }; func main() { fmt.Println(Tally{}.Add(1, 2, 3)) }",
        vec!["6"]
    ),
    variadic_named_return_empty => (
        "package main; import \"fmt\"; func size(items ...int) (n int) { n = len(items); return }; func main() { fmt.Println(size()) }",
        vec!["0"]
    ),
    variadic_passed_to_fmt_sprint => (
        "package main; import \"fmt\"; func show(parts ...interface{}) string { return fmt.Sprint(parts...) }; func main() { fmt.Println(show(\"x\", 1)) }",
        vec!["x1"]
    ),
    variadic_copy_slice_then_spread => (
        "package main; import \"fmt\"; func sum(nums ...int) int { t := 0; for _, n := range nums { t += n }; return t }; func main() { src := []int{1, 2}; dup := append([]int(nil), src...); fmt.Println(sum(dup...)) }",
        vec!["3"]
    ),
    variadic_single_arg_no_spread => (
        "package main; import \"fmt\"; func only(n ...int) int { return n[0] }; func main() { fmt.Println(only(42)) }",
        vec!["42"]
    ),
    variadic_interface_forward_to_println => (
        "package main; import \"fmt\"; func dump(parts ...interface{}) { fmt.Println(len(parts)) }; func main() { dump(true, false) }",
        vec!["2"]
    ),
    variadic_mixed_fixed_string_and_ints => (
        "package main; import \"fmt\"; func tagSum(label string, nums ...int) int { t := len(label); for _, n := range nums { t += n }; return t }; func main() { fmt.Println(tagSum(\"go\", 1, 2)) }",
        vec!["5"]
    ),
    variadic_empty_then_append_spread => (
        "package main; import \"fmt\"; func lenAfter(base []int, more ...int) int { combined := append(base, more...); return len(combined) }; func main() { fmt.Println(lenAfter([]int{1}, 2, 3)) }",
        vec!["3"]
    ),
    variadic_byte_slice_spread => (
        "package main; import \"fmt\"; func total(bytes ...byte) int { t := 0; for _, b := range bytes { t += int(b) }; return t }; func main() { data := []byte{'a', 'b'}; fmt.Println(total(data...)) }",
        vec!["195"]
    ),
}

go_compile_cases! {
    variadic_final_parameter_signature_compile =>
        "package main; func take(a int, rest ...string) int { return a + len(rest) }; func main() { _ = take(1, \"x\", \"y\") }",
    spread_requires_slice_type_compile =>
        "package main; func sink(nums ...int) int { return len(nums) }; func main() { _ = sink([]int{1, 2}...) }",
    forward_variadic_identity_compile =>
        "package main; func id(nums ...int) []int { return nums }; func wrap(nums ...int) []int { return id(nums...) }; func main() { _ = wrap(1) }",
    variadic_interface_empty_compile =>
        "package main; func pack(values ...interface{}) int { return len(values) }; func main() { _ = pack() }",
    variadic_in_interface_method_compile =>
        "package main; type Writer interface { Write(parts ...string) int }; type logger struct{}; func (l logger) Write(parts ...string) int { return len(parts) }; func main() { var w Writer = logger{}; _ = w.Write(\"a\") }",
    variadic_nested_call_spread_compile =>
        "package main; func inner(nums ...int) int { return len(nums) }; func outer(nums ...int) int { return inner(append([]int{0}, nums...)...) }; func main() { _ = outer(2, 3) }",
    variadic_func_value_call_compile =>
        "package main; func main() { fn := func(nums ...int) int { return len(nums) }; _ = fn(1, 2) }",
    variadic_returned_from_factory_compile =>
        "package main; func build() func(...int) int { return func(nums ...int) int { return len(nums) } }; func main() { f := build(); _ = f(1) }",
    variadic_with_defer_compile =>
        "package main; func log(nums ...int) int { defer func() { _ = len(nums) }(); return len(nums) }; func main() { _ = log(1, 2) }",
    variadic_range_over_parameter_compile =>
        "package main; func sum(nums ...int) int { t := 0; for _, n := range nums { t += n }; return t }; func main() { _ = sum(1, 2, 3) }",
    variadic_string_spread_compile =>
        "package main; func join(parts ...string) string { s := \"\"; for _, p := range parts { s += p }; return s }; func main() { tail := []string{\"b\"}; _ = join(append([]string{\"a\"}, tail...)...) }",
    variadic_mixed_types_separate_funcs_compile =>
        "package main; func ints(nums ...int) int { return len(nums) }; func strs(parts ...string) int { return len(parts) }; func main() { _ = ints(1); _ = strs(\"x\") }",
    variadic_assign_to_slice_var_compile =>
        "package main; func collect(nums ...int) []int { return nums }; func main() { s := collect(1, 2); _ = s[0] }",
    variadic_in_struct_literal_field_compile =>
        "package main; type cfg struct { tags []string }; func labels(parts ...string) cfg { return cfg{tags: parts} }; func main() { c := labels(\"a\", \"b\"); _ = c.tags[1] }",
    variadic_if_guard_empty_compile =>
        "package main; func ok(nums ...int) bool { return len(nums) > 0 }; func main() { if ok(1) { _ = ok() } }",
    variadic_closure_capture_compile =>
        "package main; func main() { base := 1; fn := func(nums ...int) int { return base + len(nums) }; _ = fn(2, 3) }",
    variadic_type_switch_on_interface_pack_compile =>
        "package main; func classify(values ...interface{}) int { c := 0; for _, v := range values { switch v.(type) { case int: c++ } }; return c }; func main() { _ = classify(1, \"x\", 2) }",
    variadic_spread_from_function_return_compile =>
        "package main; func batch() []int { return []int{1, 2} }; func sum(nums ...int) int { t := 0; for _, n := range nums { t += n }; return t }; func main() { _ = sum(batch()...) }",
    variadic_multiple_spread_sources_compile =>
        "package main; func merge(a ...int) int { return len(a) }; func main() { first := []int{1}; second := []int{2, 3}; _ = merge(append(first, second...)...) }",
    variadic_comma_before_spread_compile =>
        "package main; func pick(prefix int, rest ...int) int { return prefix + len(rest) }; func main() { extra := []int{4, 5}; _ = pick(1, extra...) }",
    variadic_in_select_case_compile =>
        "package main; func send(ch chan int, nums ...int) { for _, n := range nums { ch <- n } }; func main() { ch := make(chan int, 2); send(ch, 1, 2); _ = <-ch }",
    variadic_named_type_slice_spread_compile =>
        "package main; type digits []int; func sum(nums ...int) int { return len(nums) }; func main() { d := digits{1, 2}; _ = sum([]int(d)...)}",
    variadic_with_blank_import_call_compile =>
        "package main; import \"fmt\"; func show(parts ...interface{}) { _ = fmt.Sprint(parts...) }; func main() { show(1, 2) }",
    variadic_pointer_receiver_method_compile =>
        "package main; type acc struct { n int }; func (a *acc) add(nums ...int) { for _, v := range nums { a.n += v } }; func main() { x := acc{}; x.add(1, 2); _ = x.n }",
    variadic_in_comparison_compile =>
        "package main; func lenInts(nums ...int) int { return len(nums) }; func main() { _ = lenInts(1, 2) == 2 }",
    variadic_append_to_nil_from_variadic_compile =>
        "package main; func gather(nums ...int) []int { return append([]int(nil), nums...) }; func main() { _ = gather(3, 4)[0] }",
    variadic_in_for_init_compile =>
        "package main; func size(nums ...int) int { return len(nums) }; func main() { for i := size(1, 2, 3); i > 0; i-- { break } }",
    variadic_switch_on_len_compile =>
        "package main; func bucket(nums ...int) int { switch len(nums) { case 0: return 0; case 1: return 1; default: return 2 } }; func main() { _ = bucket(1, 2) }",
    variadic_map_values_spread_compile =>
        "package main; func keys(m map[string]int) int { return len(m) }; func main() { _ = keys(map[string]int{\"a\": 1}) }",
    variadic_in_return_statement_compile =>
        "package main; func max(nums ...int) int { m := nums[0]; for _, n := range nums { if n > m { m = n } }; return m }; func main() { _ = max(3, 9, 1) }",
}
