//! Multiple `init()` functions: declaration order, package-level var use, and helper calls.
//! Distinct from `test_init_blank_import.rs` (blank imports) and
//! `test_declarations_patterns.rs` (single init smoke).

go_run_cases! {
    init_two_funcs_append_letters => (
        "package main; import \"fmt\"; var trace string; func init() { trace += \"A\" }; func init() { trace += \"B\" }; func main() { fmt.Println(trace) }",
        vec!["AB"]
    ),
    init_three_funcs_build_digit_string => (
        "package main; import \"fmt\"; var digits string; func init() { digits = \"1\" }; func init() { digits += \"2\" }; func init() { digits += \"3\" }; func main() { fmt.Println(digits) }",
        vec!["123"]
    ),
    init_reads_prior_init_counter => (
        "package main; import \"fmt\"; var step int; func init() { step = 1 }; func init() { step = step * 10 }; func init() { step = step + 5 }; func main() { fmt.Println(step) }",
        vec!["15"]
    ),
    init_second_uses_first_bool_flag => (
        "package main; import \"fmt\"; var armed bool; var ready bool; func init() { armed = true }; func init() { ready = armed }; func main() { fmt.Println(ready) }",
        vec!["true"]
    ),
    init_chain_doubles_then_adds => (
        "package main; import \"fmt\"; var value int; func init() { value = 3 }; func init() { value = value * 2 }; func init() { value = value + 1 }; func main() { fmt.Println(value) }",
        vec!["7"]
    ),
    init_appends_to_shared_slice_in_order => (
        "package main; import \"fmt\"; var seq []int; func init() { seq = append(seq, 10) }; func init() { seq = append(seq, 20) }; func main() { fmt.Println(len(seq)); fmt.Println(seq[0]); fmt.Println(seq[1]) }",
        vec!["2", "10", "20"]
    ),
    init_sets_map_then_second_reads => (
        "package main; import \"fmt\"; var registry = map[string]int{}; var total int; func init() { registry[\"x\"] = 4 }; func init() { total = registry[\"x\"] + 1 }; func main() { fmt.Println(total) }",
        vec!["5"]
    ),
    init_calls_setup_then_validate => (
        "package main; import \"fmt\"; var ok bool; func setup() { ok = true }; func validate() bool { return ok }; func init() { setup() }; func init() { ok = validate() }; func main() { fmt.Println(ok) }",
        vec!["true"]
    ),
    init_helper_increments_twice => (
        "package main; import \"fmt\"; var count int; func bump() { count++ }; func init() { bump() }; func init() { bump() }; func main() { fmt.Println(count) }",
        vec!["2"]
    ),
    init_nested_helper_chain => (
        "package main; import \"fmt\"; var label string; func prefix() { label = \"go\" }; func suffix() { label += \"lang\" }; func init() { prefix() }; func init() { suffix() }; func main() { fmt.Println(label) }",
        vec!["golang"]
    ),
    init_uses_const_defined_before => (
        "package main; import \"fmt\"; const base = 8; var scaled int; func init() { scaled = base }; func init() { scaled = scaled / 2 }; func main() { fmt.Println(scaled) }",
        vec!["4"]
    ),
    init_string_builder_three_steps => (
        "package main; import \"fmt\"; var buf string; func init() { buf = \"v\" }; func init() { buf += \"y\" }; func init() { buf += \"be\" }; func main() { fmt.Println(buf) }",
        vec!["vybe"]
    ),
    init_four_funcs_numeric_pipeline => (
        "package main; import \"fmt\"; var n int; func init() { n = 2 }; func init() { n += 3 }; func init() { n *= 2 }; func init() { n -= 1 }; func main() { fmt.Println(n) }",
        vec!["9"]
    ),
    init_second_reads_slice_len_from_first => (
        "package main; import \"fmt\"; var items []string; var size int; func init() { items = []string{\"a\", \"b\"} }; func init() { size = len(items) }; func main() { fmt.Println(size) }",
        vec!["2"]
    ),
    init_assigns_struct_then_reads_field => (
        "package main; import \"fmt\"; type pair struct { a int; b int }; var p pair; func init() { p = pair{a: 3, b: 4} }; func init() { p.b = p.a + p.b }; func main() { fmt.Println(p.b) }",
        vec!["7"]
    ),
    init_loop_in_first_second_sums => (
        "package main; import \"fmt\"; var seed int; var total int; func init() { for i := 0; i < 3; i++ { seed += i } }; func init() { total = seed + 10 }; func main() { fmt.Println(total) }",
        vec!["13"]
    ),
    init_pointer_deref_chain => (
        "package main; import \"fmt\"; var target int; var ptr *int; func init() { target = 5; ptr = &target }; func init() { *ptr = *ptr + 2 }; func main() { fmt.Println(target) }",
        vec!["7"]
    ),
    init_type_alias_conversion => (
        "package main; import \"fmt\"; type score int; var high score; func init() { high = score(11) }; func init() { high = high + 1 }; func main() { fmt.Println(int(high)) }",
        vec!["12"]
    ),
    init_closure_var_capture => (
        "package main; import \"fmt\"; var fn func() int; var result int; func init() { n := 6; fn = func() int { return n } }; func init() { result = fn() }; func main() { fmt.Println(result) }",
        vec!["6"]
    ),
    init_map_two_keys_sequential => (
        "package main; import \"fmt\"; var m = map[int]string{}; func init() { m[1] = \"one\" }; func init() { m[2] = \"two\" }; func main() { fmt.Println(len(m)); fmt.Println(m[2]) }",
        vec!["2", "two"]
    ),
    init_bool_toggle_twice => (
        "package main; import \"fmt\"; var flag bool; func init() { flag = !flag }; func init() { flag = !flag }; func main() { fmt.Println(flag) }",
        vec!["false"]
    ),
    init_array_element_mutation => (
        "package main; import \"fmt\"; var arr = [3]int{1, 1, 1}; func init() { arr[0] = 2 }; func init() { arr[1] = arr[0] + 1 }; func main() { fmt.Println(arr[0]); fmt.Println(arr[1]) }",
        vec!["2", "3"]
    ),
    init_three_call_same_helper => (
        "package main; import \"fmt\"; var tally int; func add(n int) { tally += n }; func init() { add(1) }; func init() { add(2) }; func init() { add(3) }; func main() { fmt.Println(tally) }",
        vec!["6"]
    ),
    init_uses_iota_const_group => (
        "package main; import \"fmt\"; const ( First = iota; Second; Third ); var picked int; func init() { picked = Second }; func init() { picked = picked + Third }; func main() { fmt.Println(picked) }",
        vec!["3"]
    ),
    init_rune_accumulation => (
        "package main; import \"fmt\"; var ch rune; func init() { ch = 'A' }; func init() { ch = ch + 1 }; func main() { fmt.Println(int(ch)) }",
        vec!["66"]
    ),
    init_float64_product => (
        "package main; import \"fmt\"; var ratio float64; func init() { ratio = 2.0 }; func init() { ratio = ratio * 1.5 }; func main() { fmt.Println(ratio) }",
        vec!["3"]
    ),
    init_byte_slice_append => (
        "package main; import \"fmt\"; var data []byte; func init() { data = append(data, 'x') }; func init() { data = append(data, 'y') }; func main() { fmt.Println(len(data)); fmt.Println(int(data[1])) }",
        vec!["2", "121"]
    ),
    init_named_return_helper => (
        "package main; import \"fmt\"; var stored int; func read() (n int) { n = 9; return }; func init() { stored = read() }; func init() { stored++ }; func main() { fmt.Println(stored) }",
        vec!["10"]
    ),
    init_interface_value_assignment => (
        "package main; import \"fmt\"; var holder interface{}; var tag string; func init() { holder = 42 }; func init() { tag = fmt.Sprint(holder) }; func main() { fmt.Println(tag) }",
        vec!["42"]
    ),
    init_deferred_side_effect_via_helper => (
        "package main; import \"fmt\"; var order string; func mark(c string) { order += c }; func init() { mark(\"1\"); defer mark(\"d1\") }; func init() { mark(\"2\"); defer mark(\"d2\") }; func main() { fmt.Println(order) }",
        vec!["12"]
    ),
}

go_compile_cases! {
    five_init_functions_sequential_compile =>
        "package main; var phase int; func init() { phase = 1 }; func init() { phase++ }; func init() { phase++ }; func init() { phase++ }; func init() { phase++ }; func main() { _ = phase }",
    init_reads_package_var_from_prior_init_compile =>
        "package main; var a int; var b int; func init() { a = 7 }; func init() { b = a + 1 }; func main() { _ = b }",
    init_calls_two_package_helpers_compile =>
        "package main; var ready bool; func arm() { ready = true }; func confirm() { _ = ready }; func init() { arm() }; func init() { confirm() }; func main() {}",
    init_with_range_loop_compile =>
        "package main; var sum int; func init() { for i := range 4 { sum += i } }; func main() { _ = sum }",
    init_with_select_on_channel_compile =>
        "package main; func init() { ch := make(chan int, 1); ch <- 1; select { case v := <-ch: _ = v default: } }; func main() {}",
    init_assigns_func_value_then_calls_compile =>
        "package main; var double func(int) int; func init() { double = func(n int) int { return n * 2 } }; func init() { _ = double(3) }; func main() {}",
    init_mutates_struct_pointer_compile =>
        "package main; type node struct { next *node; val int }; var head node; func init() { head.val = 1; head.next = &head }; func main() { _ = head.next.val }",
    init_uses_imported_fmt_sprint_compile =>
        "package main; import \"fmt\"; var s string; func init() { s = fmt.Sprint(1, 2) }; func main() { _ = s }",
    init_after_blank_import_strings_compile =>
        "package main; import _ \"strings\"; var seeded int; func init() { seeded = 1 }; func init() { seeded++ }; func main() { _ = seeded }",
    init_with_type_switch_compile =>
        "package main; var tag string; func init() { var v interface{} = 3; switch v.(type) { case int: tag = \"int\" default: tag = \"other\" } }; func main() { _ = tag }",
    init_with_labeled_break_compile =>
        "package main; var count int; func init() { outer: for i := 0; i < 5; i++ { count++; break outer } }; func main() { _ = count }",
    init_nested_anonymous_func_compile =>
        "package main; var n int; func init() { func() { n = 4 }() }; func main() { _ = n }",
    init_writes_to_map_of_slices_compile =>
        "package main; var groups = map[string][]int{}; func init() { groups[\"a\"] = []int{1} }; func init() { groups[\"a\"] = append(groups[\"a\"], 2) }; func main() { _ = groups }",
    init_uses_const_arithmetic_compile =>
        "package main; const ( A = 2; B = A * 3 ); var c int; func init() { c = B + 1 }; func main() { _ = c }",
    init_method_on_named_type_compile =>
        "package main; type counter int; func (c *counter) inc() { *c++ }; var total counter; func init() { total.inc() }; func main() { _ = total }",
    init_embedded_struct_field_compile =>
        "package main; type base struct { x int }; type child struct { base; y int }; var c child; func init() { c.x = 1; c.y = 2 }; func main() { _ = c.x + c.y }",
    init_with_if_init_short_decl_compile =>
        "package main; var ok bool; func init() { if n := 3; n > 0 { ok = true } }; func main() { _ = ok }",
    init_for_range_over_string_compile =>
        "package main; var count int; func init() { for range \"go\" { count++ } }; func main() { _ = count }",
    init_slice_reslice_compile =>
        "package main; var part []int; func init() { all := []int{1, 2, 3, 4}; part = all[1:3] }; func main() { _ = part[0] }",
    init_map_composite_literal_compile =>
        "package main; var table map[string]int; func init() { table = map[string]int{\"k\": 9} }; func main() { _ = table[\"k\"] }",
    init_channel_make_buffered_compile =>
        "package main; var ch chan int; func init() { ch = make(chan int, 2); ch <- 1 }; func main() { _ = <-ch }",
    init_defer_in_init_body_compile =>
        "package main; var x int; func init() { defer func() { x++ }(); x = 1 }; func main() { _ = x }",
    init_two_inits_both_use_package_func_var_compile =>
        "package main; var apply func(int) int; func init() { apply = func(n int) int { return n + 1 } }; func init() { _ = apply(2) }; func main() {}",
    init_with_comma_ok_map_lookup_compile =>
        "package main; var m = map[string]int{\"a\": 1}; var v int; func init() { v, _ = m[\"a\"] }; func main() { _ = v }",
    init_grouped_var_block_used_compile =>
        "package main; var ( x int; y int ); func init() { x = 1; y = 2 }; func init() { x = x + y }; func main() { _ = x }",
}
