//! Extended blank identifier (`_`) usage: side-effect imports, range discard,
//! multi-assign discard, struct literal omission, interface stubs, and
//! anonymous struct fields.
//! Distinct from `test_init_blank_import.rs` and `test_declarations_patterns.rs`.

use crate::helpers::*;

go_run_cases! {
    blank_import_side_effect_with_local_init => (
        "package main; import \"fmt\"; import _ \"strings\"; var ready int; func init() { ready = 1 }; func main() { fmt.Println(ready) }",
        vec!["1"]
    ),
    blank_range_discard_index_only => (
        "package main; import \"fmt\"; func main() { total := 0; for _, v := range []int{2, 3, 4} { total += v }; fmt.Println(total) }",
        vec!["9"]
    ),
    blank_range_discard_value_count_index => (
        "package main; import \"fmt\"; func main() { count := 0; for range []string{\"a\", \"b\", \"c\"} { count++ }; fmt.Println(count) }",
        vec!["3"]
    ),
    blank_range_map_discard_value => (
        "package main; import \"fmt\"; func main() { keys := 0; for k := range map[string]int{\"x\": 1, \"y\": 2} { keys += len(k) }; fmt.Println(keys) }",
        vec!["2"]
    ),
    blank_range_string_discard_index => (
        "package main; import \"fmt\"; func main() { total := 0; for _, r := range \"ab\" { total += int(r) }; fmt.Println(total) }",
        vec!["195"]
    ),
    blank_multi_assign_keep_first => (
        "package main; import \"fmt\"; func pair() (int, string) { return 7, \"go\" }; func main() { a, _ := pair(); fmt.Println(a) }",
        vec!["7"]
    ),
    blank_multi_assign_keep_second => (
        "package main; import \"fmt\"; func pair() (int, string) { return 3, \"vybe\" }; func main() { _, b := pair(); fmt.Println(b) }",
        vec!["vybe"]
    ),
    blank_multi_assign_three_returns => (
        "package main; import \"fmt\"; func triple() (int, int, int) { return 1, 2, 3 }; func main() { _, mid, _ := triple(); fmt.Println(mid) }",
        vec!["2"]
    ),
    blank_multi_assign_map_ok => (
        "package main; import \"fmt\"; func main() { m := map[string]int{\"k\": 9}; v, _ := m[\"k\"]; fmt.Println(v) }",
        vec!["9"]
    ),
    blank_multi_assign_map_missing_ok => (
        "package main; import \"fmt\"; func main() { m := map[string]int{}; _, ok := m[\"missing\"]; fmt.Println(ok) }",
        vec!["false"]
    ),
    blank_struct_literal_omit_unexported_field => (
        "package main; import \"fmt\"; type point struct { x int; y int }; func main() { p := point{x: 3}; fmt.Println(p.x); fmt.Println(p.y) }",
        vec!["3", "0"]
    ),
    blank_struct_literal_named_partial => (
        "package main; import \"fmt\"; type cfg struct { host string; port int }; func main() { c := cfg{port: 8080}; fmt.Println(c.port) }",
        vec!["8080"]
    ),
    blank_discard_append_return => (
        "package main; import \"fmt\"; func main() { s := []int{1}; _ = append(s, 2); fmt.Println(len(s)) }",
        vec!["1"]
    ),
    blank_discard_arithmetic_side_effect => (
        "package main; import \"fmt\"; func main() { x := 5; _ = x + 3; fmt.Println(x) }",
        vec!["5"]
    ),
    blank_discard_type_assertion => (
        "package main; import \"fmt\"; func main() { var v interface{} = 42; _, ok := v.(int); fmt.Println(ok) }",
        vec!["true"]
    ),
    blank_discard_channel_send => (
        "package main; import \"fmt\"; func main() { ch := make(chan int, 1); ch <- 1; _ = <-ch; fmt.Println(len(ch)) }",
        vec!["0"]
    ),
    blank_for_init_discard => (
        "package main; import \"fmt\"; func main() { total := 0; for _ = range 3 { total++ }; fmt.Println(total) }",
        vec!["3"]
    ),
    blank_range_int_discard_index => (
        "package main; import \"fmt\"; func main() { count := 0; for range 4 { count++ }; fmt.Println(count) }",
        vec!["4"]
    ),
    blank_swap_via_multi_assign => (
        "package main; import \"fmt\"; func main() { a, b := 1, 2; a, b = b, a; _, _ = a, b; fmt.Println(a); fmt.Println(b) }",
        vec!["2", "1"]
    ),
    blank_discard_func_call_returns => (
        "package main; import \"fmt\"; func divmod(a int, b int) (int, int) { return a / b, a % b }; func main() { q, _ := divmod(10, 3); fmt.Println(q) }",
        vec!["3"]
    ),
    blank_discard_slice_index => (
        "package main; import \"fmt\"; func main() { s := []int{10, 20, 30}; _, last := s[0], s[2]; fmt.Println(last) }",
        vec!["30"]
    ),
    blank_import_fmt_still_usable => (
        "package main; import \"fmt\"; import _ \"strings\"; func main() { fmt.Println(\"ok\") }",
        vec!["ok"]
    ),
    blank_discard_defer_call => (
        "package main; import \"fmt\"; func main() { x := 1; defer func() { _ = x }(); fmt.Println(x) }",
        vec!["1"]
    ),
    blank_discard_comma_ok_len => (
        "package main; import \"fmt\"; func main() { s := []int{1, 2, 3}; _ = len(s); fmt.Println(s[0]) }",
        vec!["1"]
    ),
    blank_range_nested_outer_discard => (
        "package main; import \"fmt\"; func main() { total := 0; for _, row := range [][]int{{1, 2}, {3}} { for _, v := range row { total += v } }; fmt.Println(total) }",
        vec!["6"]
    ),
}

go_compile_cases! {
    blank_import_math_side_effect_compile =>
        "package main; import _ \"math\"; func main() {}",
    blank_import_grouped_encoding_compile =>
        "package main; import ( _ \"encoding/json\"; _ \"encoding/hex\" ); func main() {}",
    blank_import_with_named_fmt_compile =>
        "package main; import ( \"fmt\"; _ \"strings\" ); func main() { _ = fmt.Sprint(1) }",
    blank_range_slice_index_discard_compile =>
        "package main; func main() { for _, v := range []int{1, 2} { _ = v } }",
    blank_range_map_key_discard_compile =>
        "package main; func main() { for _, v := range map[int]string{1: \"a\"} { _ = v } }",
    blank_range_channel_discard_compile =>
        "package main; func main() { ch := make(chan int, 1); ch <- 1; close(ch); for range ch { } }",
    blank_multi_assign_swap_compile =>
        "package main; func main() { a, b := 1, 2; a, b = b, a; _, _ = a, b }",
    blank_multi_assign_func_returns_compile =>
        "package main; func duo() (int, int) { return 1, 2 }; func main() { _, y := duo(); _ = y }",
    blank_struct_literal_omit_fields_compile =>
        "package main; type node struct { id int; name string }; func main() { n := node{id: 1}; _ = n.id }",
    blank_anonymous_struct_field_compile =>
        "package main; func main() { type inner struct { _ int; v int }; x := inner{v: 3}; _ = x.v }",
    blank_interface_method_impl_discard_compile =>
        "package main; type Closer interface { Close() error }; type resource struct{}; func (r resource) Close() error { return nil }; func main() { var c Closer = resource{}; _ = c }",
    blank_interface_embedded_method_compile =>
        "package main; type Reader interface { Read(p []byte) (n int, err error) }; type rw struct{}; func (r rw) Read(p []byte) (int, error) { return 0, nil }; func main() { var rd Reader = rw{}; _ = rd }",
    blank_type_switch_discard_compile =>
        "package main; func main() { var v interface{} = \"x\"; switch v.(type) { case string: _ = v } }",
    blank_for_init_short_decl_compile =>
        "package main; func main() { for _ = range 2 { } }",
    blank_discard_map_delete_compile =>
        "package main; func main() { m := map[string]int{\"a\": 1}; delete(m, \"a\"); _, ok := m[\"a\"]; _ = ok }",
    blank_discard_select_receive_compile =>
        "package main; func main() { ch := make(chan int, 1); ch <- 1; select { case _ = <-ch: } }",
    blank_discard_type_conversion_compile =>
        "package main; type score int; func main() { var s score = 3; _ = int(s) }",
    blank_discard_composite_array_literal_compile =>
        "package main; func main() { _ = [2]int{1, 2}; a := [2]int{1}; _ = a[0] }",
    blank_discard_composite_map_literal_compile =>
        "package main; func main() { _ = map[string]int{\"k\": 1}; m := map[string]int{}; _ = m }",
    blank_discard_slice_expression_compile =>
        "package main; func main() { s := []int{1, 2, 3}; _ = s[1:2]; _ = s[0] }",
    blank_discard_func_literal_call_compile =>
        "package main; func main() { _ = func(x int) int { return x }(4) }",
    blank_discard_method_call_compile =>
        "package main; type T struct{}; func (t T) M() int { return 1 }; func main() { _ = T{}.M() }",
    blank_discard_unary_operators_compile =>
        "package main; func main() { x := 3; _ = -x; _ = ^x }",
    blank_discard_binary_operators_compile =>
        "package main; func main() { _, _ = 1+2, 3*4 }",
    blank_discard_comparison_compile =>
        "package main; func main() { _ = 1 < 2 }",
    blank_discard_address_of_compile =>
        "package main; func main() { x := 1; _ = &x }",
    blank_discard_star_deref_compile =>
        "package main; func main() { x := 1; p := &x; _ = *p }",
    blank_discard_range_over_int_compile =>
        "package main; func main() { for range 2 { _ = 1 } }",
    blank_discard_import_underscore_alias_compile =>
        "package main; import f \"fmt\"; import _ \"strings\"; func main() { _ = f.Sprint(\"x\") }",
    blank_anonymous_struct_literal_embed_compile =>
        "package main; func main() { type outer struct { _ struct{}; n int }; o := outer{n: 2}; _ = o.n }",
}
