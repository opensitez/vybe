//! Type assertions: comma-ok, wrong-type panic/recover, nil interface semantics,
//! assertion to interface types. Distinct from `test_interface_nil_comparable.rs`
//! (equality/comparable) and `test_type_switch_extended.rs` (type switches).

use crate::helpers::*;

go_run_cases! {
    comma_ok_assert_int_from_empty_interface_true =>
        ("package main; import \"fmt\"; func main() { var v interface{} = 42; n, ok := v.(int); fmt.Println(n); fmt.Println(ok) }", vec!["42", "true"]),
    comma_ok_assert_int_from_empty_interface_false =>
        ("package main; import \"fmt\"; func main() { var v interface{} = \"x\"; _, ok := v.(int); fmt.Println(ok) }", vec!["false"]),
    comma_ok_assert_string_from_empty_interface_true =>
        ("package main; import \"fmt\"; func main() { var v interface{} = \"go\"; s, ok := v.(string); fmt.Println(s); fmt.Println(ok) }", vec!["go", "true"]),
    comma_ok_assert_struct_from_interface_true =>
        ("package main; import \"fmt\"; type point struct { x int; y int }; func main() { var v interface{} = point{x: 1, y: 2}; p, ok := v.(point); fmt.Println(p.x + p.y); fmt.Println(ok) }", vec!["3", "true"]),
    comma_ok_assert_pointer_from_interface_true =>
        ("package main; import \"fmt\"; type node struct { id int }; func main() { n := &node{id: 5}; var v interface{} = n; p, ok := v.(*node); fmt.Println(p.id); fmt.Println(ok) }", vec!["5", "true"]),
    comma_ok_assert_pointer_from_interface_false =>
        ("package main; import \"fmt\"; type node struct { id int }; func main() { var v interface{} = node{id: 1}; _, ok := v.(*node); fmt.Println(ok) }", vec!["false"]),
    comma_ok_assert_interface_from_interface_true =>
        ("package main; import \"fmt\"; type reader interface { read() int }; type book struct { pages int }; func (b book) read() int { return b.pages }; func main() { var concrete reader = book{pages: 10}; var v interface{} = concrete; r, ok := v.(reader); fmt.Println(r.read()); fmt.Println(ok) }", vec!["10", "true"]),
    comma_ok_assert_named_interface_from_empty_false =>
        ("package main; import \"fmt\"; type speaker interface { talk() string }; func main() { var v interface{} = 1; _, ok := v.(speaker); fmt.Println(ok) }", vec!["false"]),
    single_assert_int_from_interface_runtime =>
        ("package main; import \"fmt\"; func main() { var v interface{} = 7; fmt.Println(v.(int)) }", vec!["7"]),
    wrong_type_assertion_panic_recovered_runtime =>
        ("package main; import \"fmt\"; func main() { defer func() { fmt.Println(recover() != nil) }(); var v interface{} = 1; _ = v.(string) }", vec!["true"]),
    typed_nil_pointer_in_interface_not_equal_nil =>
        ("package main; import \"fmt\"; func main() { var p *int; var v interface{} = p; fmt.Println(v == nil) }", vec!["false"]),
    typed_nil_pointer_in_named_interface_not_nil =>
        ("package main; import \"fmt\"; type holder interface { size() int }; type box struct { n int }; func (b *box) size() int { return b.n }; func main() { var p *box; var h holder = p; fmt.Println(h == nil) }", vec!["false"]),
    untyped_nil_interface_is_nil =>
        ("package main; import \"fmt\"; func main() { var v interface{}; fmt.Println(v == nil) }", vec!["true"]),
    untyped_nil_assigned_to_interface_stays_nil =>
        ("package main; import \"fmt\"; func main() { var v interface{} = nil; fmt.Println(v == nil) }", vec!["true"]),
    typed_nil_slice_in_interface_not_nil =>
        ("package main; import \"fmt\"; func main() { var s []int; var v interface{} = s; fmt.Println(v == nil) }", vec!["false"]),
    typed_nil_map_in_interface_not_nil =>
        ("package main; import \"fmt\"; func main() { var m map[string]int; var v interface{} = m; fmt.Println(v == nil) }", vec!["false"]),
    typed_nil_func_in_interface_not_nil =>
        ("package main; import \"fmt\"; func main() { var f func(); var v interface{} = f; fmt.Println(v == nil) }", vec!["false"]),
    typed_nil_channel_in_interface_not_nil =>
        ("package main; import \"fmt\"; func main() { var ch chan int; var v interface{} = ch; fmt.Println(v == nil) }", vec!["false"]),
    assert_to_error_interface_from_typed_nil =>
        ("package main; import \"fmt\"; type myErr struct { msg string }; func (e *myErr) Error() string { return e.msg }; func main() { var p *myErr; var err error = p; fmt.Println(err == nil); _, ok := err.(*myErr); fmt.Println(ok) }", vec!["false", "true"]),
    assert_interface_type_from_empty_interface =>
        ("package main; import \"fmt\"; type fmtStringer interface { String() string }; type myInt int; func (m myInt) String() string { return \"n\" }; func main() { var v interface{} = myInt(3); s, ok := v.(fmtStringer); fmt.Println(s.String()); fmt.Println(ok) }", vec!["n", "true"]),
    assert_empty_interface_from_named_interface =>
        ("package main; import \"fmt\"; type speaker interface { talk() string }; type bot struct{}; func (bot) talk() string { return \"hi\" }; func main() { var s speaker = bot{}; var v interface{} = s; b, ok := v.(bot); fmt.Println(b.talk()); fmt.Println(ok) }", vec!["hi", "true"]),
    comma_ok_on_nil_interface_value =>
        ("package main; import \"fmt\"; func main() { var v interface{}; _, ok := v.(int); fmt.Println(ok) }", vec!["false"]),
    assert_slice_type_from_interface =>
        ("package main; import \"fmt\"; func main() { var v interface{} = []int{1, 2}; s, ok := v.([]int); fmt.Println(len(s)); fmt.Println(ok) }", vec!["2", "true"]),
    assert_map_type_from_interface =>
        ("package main; import \"fmt\"; func main() { var v interface{} = map[string]int{\"a\": 1}; m, ok := v.(map[string]int); fmt.Println(m[\"a\"]); fmt.Println(ok) }", vec!["1", "true"]),
    assert_func_type_from_interface =>
        ("package main; import \"fmt\"; func main() { fn := func(x int) int { return x + 1 }; var v interface{} = fn; f, ok := v.(func(int) int); fmt.Println(f(4)); fmt.Println(ok) }", vec!["5", "true"]),
    assert_array_type_from_interface =>
        ("package main; import \"fmt\"; func main() { var v interface{} = [2]int{3, 4}; a, ok := v.([2]int); fmt.Println(a[0] + a[1]); fmt.Println(ok) }", vec!["7", "true"]),
    assert_bool_from_interface =>
        ("package main; import \"fmt\"; func main() { var v interface{} = true; b, ok := v.(bool); fmt.Println(b); fmt.Println(ok) }", vec!["true", "true"]),
    assert_float64_from_interface =>
        ("package main; import \"fmt\"; func main() { var v interface{} = 2.5; f, ok := v.(float64); fmt.Println(f); fmt.Println(ok) }", vec!["2.5", "true"]),
    assert_byte_from_interface =>
        ("package main; import \"fmt\"; func main() { var v interface{} = byte(65); b, ok := v.(byte); fmt.Println(b); fmt.Println(ok) }", vec!["65", "true"]),
    assert_rune_from_interface =>
        ("package main; import \"fmt\"; func main() { var v interface{} = rune(8364); r, ok := v.(rune); fmt.Println(int(r)); fmt.Println(ok) }", vec!["8364", "true"]),
    assert_named_type_from_interface =>
        ("package main; import \"fmt\"; type counter int; func main() { var v interface{} = counter(9); c, ok := v.(counter); fmt.Println(int(c)); fmt.Println(ok) }", vec!["9", "true"]),
    assert_pointer_struct_deref_after_ok =>
        ("package main; import \"fmt\"; type pair struct { a int }; func main() { p := &pair{a: 6}; var v interface{} = p; if q, ok := v.(*pair); ok { fmt.Println(q.a) } else { fmt.Println(0) } }", vec!["6"]),
    assert_interface_implementer_method_call =>
        ("package main; import \"fmt\"; type greeter interface { greet() string }; type hi struct{}; func (hi) greet() string { return \"yo\" }; func main() { var v interface{} = hi{}; if g, ok := v.(greeter); ok { fmt.Println(g.greet()) } else { fmt.Println(\"no\") } }", vec!["yo"]),
    typed_nil_assert_to_concrete_pointer_ok =>
        ("package main; import \"fmt\"; type widget struct { n int }; func main() { var p *widget; var v interface{} = p; _, ok := v.(*widget); fmt.Println(ok) }", vec!["true"]),
    interface_nil_vs_typed_nil_equality =>
        ("package main; import \"fmt\"; func main() { var empty interface{}; var p *int; var typed interface{} = p; fmt.Println(empty == typed) }", vec!["false"]),
    two_typed_nil_same_type_equal_in_interface =>
        ("package main; import \"fmt\"; func main() { var a *int; var b *int; var left interface{} = a; var right interface{} = b; fmt.Println(left == right) }", vec!["true"]),
    two_typed_nil_different_types_not_equal =>
        ("package main; import \"fmt\"; func main() { var pi *int; var ps *string; var left interface{} = pi; var right interface{} = ps; fmt.Println(left == right) }", vec!["false"]),
    assert_chain_comma_ok_in_conditional =>
        ("package main; import \"fmt\"; func pick(v interface{}) { if s, ok := v.(string); ok { fmt.Println(s) } else if n, ok := v.(int); ok { fmt.Println(n) } else { fmt.Println(\"none\") } }; func main() { pick(3) }", vec!["3"]),
    recover_after_assert_panic_allows_continue =>
        ("package main; import \"fmt\"; func main() { caught := false; func() { defer func() { if recover() != nil { caught = true } }(); var v interface{} = 1; _ = v.(bool) }(); if caught { fmt.Println(\"ok\") } else { fmt.Println(\"no\") } }", vec!["ok"]),
    assert_uint_from_interface =>
        ("package main; import \"fmt\"; func main() { var v interface{} = uint(12); u, ok := v.(uint); fmt.Println(u); fmt.Println(ok) }", vec!["12", "true"]),
    assert_int32_from_interface =>
        ("package main; import \"fmt\"; func main() { var v interface{} = int32(-3); n, ok := v.(int32); fmt.Println(n); fmt.Println(ok) }", vec!["-3", "true"]),
    assert_complex128_from_interface =>
        ("package main; import \"fmt\"; func main() { var v interface{} = complex(1, 2); c, ok := v.(complex128); fmt.Println(real(c) + imag(c)); fmt.Println(ok) }", vec!["3", "true"]),
    assert_chan_from_interface =>
        ("package main; import \"fmt\"; func main() { ch := make(chan int, 1); var v interface{} = ch; c, ok := v.(chan int); c <- 5; fmt.Println(<-c); fmt.Println(ok) }", vec!["5", "true"]),
    wrong_type_assert_to_struct_recovered =>
        ("package main; import \"fmt\"; type a struct{}; type b struct{}; func main() { defer func() { fmt.Println(recover() != nil) }(); var v interface{} = a{}; _ = v.(b) }", vec!["true"]),
    comma_ok_assert_error_from_interface =>
        ("package main; import \"fmt\"; import \"errors\"; func main() { var v interface{} = errors.New(\"e\"); e, ok := v.(error); fmt.Println(e.Error()); fmt.Println(ok) }", vec!["e", "true"]),
    assert_pointer_value_mismatch_comma_ok_false =>
        ("package main; import \"fmt\"; type node struct { v int }; func main() { n := node{v: 2}; var v interface{} = n; _, ok := v.(*node); fmt.Println(ok) }", vec!["false"]),
    assert_value_from_pointer_boxed_ok =>
        ("package main; import \"fmt\"; type node struct { v int }; func main() { n := &node{v: 2}; var v interface{} = n; p, ok := v.(*node); fmt.Println(p.v); fmt.Println(ok) }", vec!["2", "true"]),
    typed_nil_error_interface_not_nil =>
        ("package main; import \"fmt\"; func main() { var err error; fmt.Println(err == nil) }", vec!["true"]),
    typed_nil_custom_error_in_error_interface =>
        ("package main; import \"fmt\"; type e struct { msg string }; func (err *e) Error() string { return err.msg }; func main() { var p *e; var err error = p; fmt.Println(err == nil) }", vec!["false"]),
}

go_compile_cases! {
    type_assertion_to_concrete_compile =>
        "package main; func main() { var v interface{} = 1; _ = v.(int) }",
    type_assertion_comma_ok_compile =>
        "package main; func main() { var v interface{} = 1; _, ok := v.(int); _ = ok }",
    type_assertion_to_interface_type_compile =>
        "package main; type reader interface { read() int }; func main() { var v interface{} = 1; _, ok := v.(reader); _ = ok }",
    type_assertion_to_pointer_compile =>
        "package main; type node struct{}; func main() { var v interface{} = &node{}; _ = v.(*node) }",
    type_assertion_on_named_interface_field_compile =>
        "package main; type speaker interface { talk() string }; type holder struct { value interface{} }; func main() { h := holder{value: 1}; _, _ = h.value.(int) }",
    type_assertion_result_used_in_switch_compile =>
        "package main; func main() { var v interface{} = \"x\"; switch s := v.(type) { case string: _ = s } }",
    assert_error_interface_from_concrete_compile =>
        "package main; import \"errors\"; func main() { err := errors.New(\"x\"); _, ok := err.(error); _ = ok }",
    assert_to_any_alias_compile =>
        "package main; func main() { var v any = 1; _, ok := v.(int); _ = ok }",
    typed_nil_in_named_interface_param_compile =>
        "package main; type worker interface { work() }; type task struct{}; func (t *task) work() {}; func accept(w worker) bool { return w == nil }; func main() { var p *task; _ = accept(p) }",
    interface_assertion_in_assignment_compile =>
        "package main; type fmtStringer interface { String() string }; func main() { var v interface{} = 1; _, ok := v.(fmtStringer); _ = ok }",
}

macro_rules! go_compile_fail_cases {
    ($($name:ident => $src:expr,)+) => {
        $(#[test] fn $name() { assert!(!compile_ok_check($src)); })+
    };
}

go_compile_fail_cases! {
    assert_concrete_int_to_string_compile_fail =>
        "package main; func main() { var x int = 1; _ = x.(string) }",
    assert_non_interface_to_concrete_compile_fail =>
        "package main; type a struct{}; func main() { var x a; _ = x.(a) }",
    assert_interface_impossible_type_compile_fail =>
        "package main; type reader interface { read() int }; type writer interface { write(int) }; func main() { var r reader; _ = r.(writer) }",
    type_assertion_on_non_interface_expr_compile_fail =>
        "package main; func main() { x := 1; _ = x.(int) }",
    comma_ok_on_non_interface_compile_fail =>
        "package main; func main() { x := 1; _, ok := x.(int); _ = ok }",
}
