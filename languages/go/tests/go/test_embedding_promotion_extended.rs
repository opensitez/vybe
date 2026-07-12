//! Struct embedding: field/method promotion, name collisions, pointer vs value embeds,
//! two-level nesting. Distinct from `test_struct_embedding_advanced.rs`,
//! `test_struct_embedding_extra.rs`, and `test_lang_interfaces_embedding.rs`.

use crate::helpers::*;

go_run_cases! {
    promoted_int_field_read_runtime =>
        ("package main; import \"fmt\"; type inner struct { count int }; type outer struct { inner }; func main() { o := outer{inner: inner{count: 12}}; fmt.Println(o.count) }", vec!["12"]),
    promoted_int_field_write_runtime =>
        ("package main; import \"fmt\"; type inner struct { count int }; type outer struct { inner }; func main() { o := outer{inner: inner{count: 1}}; o.count = 9; fmt.Println(o.count) }", vec!["9"]),
    promoted_method_no_args_runtime =>
        ("package main; import \"fmt\"; type inner struct{}; func (inner) tag() string { return \"in\" }; type outer struct { inner }; func main() { fmt.Println(outer{}.tag()) }", vec!["in"]),
    promoted_method_with_args_runtime =>
        ("package main; import \"fmt\"; type inner struct { base int }; func (i inner) add(d int) int { return i.base + d }; type outer struct { inner }; func main() { o := outer{inner: inner{base: 3}}; fmt.Println(o.add(4)) }", vec!["7"]),
    outer_field_shadows_promoted_name_runtime =>
        ("package main; import \"fmt\"; type inner struct { name string }; type outer struct { inner; name string }; func main() { o := outer{inner: inner{name: \"in\"}, name: \"out\"}; fmt.Println(o.name); fmt.Println(o.inner.name) }", vec!["out", "in"]),
    explicit_embedded_type_field_access_runtime =>
        ("package main; import \"fmt\"; type inner struct { name string }; type outer struct { inner; name string }; func main() { o := outer{inner: inner{name: \"in\"}, name: \"out\"}; fmt.Println(o.inner.name) }", vec!["in"]),
    outer_method_shadows_promoted_method_runtime =>
        ("package main; import \"fmt\"; type inner struct{}; func (inner) label() string { return \"inner\" }; type outer struct { inner }; func (outer) label() string { return \"outer\" }; func main() { fmt.Println(outer{}.label()) }", vec!["outer"]),
    explicit_embedded_type_method_call_runtime =>
        ("package main; import \"fmt\"; type inner struct{}; func (inner) label() string { return \"inner\" }; type outer struct { inner }; func (outer) label() string { return \"outer\" }; func main() { o := outer{}; fmt.Println(o.inner.label()) }", vec!["inner"]),
    value_embed_value_field_promotion_runtime =>
        ("package main; import \"fmt\"; type a struct { x int }; type b struct { a }; func main() { v := b{a: a{x: 5}}; fmt.Println(v.x) }", vec!["5"]),
    value_embed_pointer_field_promotion_runtime =>
        ("package main; import \"fmt\"; type a struct { x int }; type b struct { *a }; func main() { v := b{a: &a{x: 6}}; fmt.Println(v.x) }", vec!["6"]),
    value_embed_nil_pointer_field_zero_runtime =>
        ("package main; import \"fmt\"; type a struct { x int }; type b struct { *a }; func main() { var v b; fmt.Println(v.a == nil) }", vec!["true"]),
    pointer_receiver_on_value_embedded_promoted_runtime =>
        ("package main; import \"fmt\"; type inner struct { n int }; func (i *inner) double() { i.n *= 2 }; type outer struct { inner }; func main() { o := outer{inner: inner{n: 3}}; o.double(); fmt.Println(o.n) }", vec!["6"]),
    two_level_field_promotion_runtime =>
        ("package main; import \"fmt\"; type leaf struct { val int }; type branch struct { leaf }; type trunk struct { branch }; func main() { t := trunk{branch: branch{leaf: leaf{val: 7}}}; fmt.Println(t.val) }", vec!["7"]),
    two_level_method_promotion_runtime =>
        ("package main; import \"fmt\"; type leaf struct{}; func (leaf) deep() string { return \"L\" }; type branch struct { leaf }; type trunk struct { branch }; func main() { fmt.Println(trunk{}.deep()) }", vec!["L"]),
    two_level_explicit_middle_access_runtime =>
        ("package main; import \"fmt\"; type leaf struct { val int }; type branch struct { leaf }; type trunk struct { branch }; func main() { t := trunk{branch: branch{leaf: leaf{val: 9}}}; fmt.Println(t.branch.leaf.val) }", vec!["9"]),
    dual_embedded_distinct_fields_runtime =>
        ("package main; import \"fmt\"; type axis struct { x int }; type ord struct { y int }; type point struct { axis; ord }; func main() { p := point{axis: axis{x: 2}, ord: ord{y: 5}}; fmt.Println(p.x + p.y) }", vec!["7"]),
    dual_embedded_distinct_methods_runtime =>
        ("package main; import \"fmt\"; type north struct{}; func (north) letter() string { return \"N\" }; type east struct{}; func (east) letter() string { return \"E\" }; type compass struct { north; east }; func main() { c := compass{}; fmt.Println(c.north.letter()); fmt.Println(c.east.letter()) }", vec!["N", "E"]),
    triple_embedded_field_chain_runtime =>
        ("package main; import \"fmt\"; type c struct { n int }; type b struct { c }; type a struct { b }; func main() { v := a{b: b{c: c{n: 11}}}; fmt.Println(v.n) }", vec!["11"]),
    embedded_string_field_concat_runtime =>
        ("package main; import \"fmt\"; type inner struct { left string; right string }; type outer struct { inner }; func main() { o := outer{inner: inner{left: \"go\", right: \"lang\"}}; fmt.Println(o.left + o.right) }", vec!["golang"]),
    embedded_bool_in_condition_runtime =>
        ("package main; import \"fmt\"; type inner struct { ok bool }; type outer struct { inner }; func main() { o := outer{inner: inner{ok: true}}; if o.ok { fmt.Println(1) } else { fmt.Println(0) } }", vec!["1"]),
    promoted_field_after_struct_copy_runtime =>
        ("package main; import \"fmt\"; type inner struct { n int }; type outer struct { inner }; func main() { a := outer{inner: inner{n: 2}}; b := a; b.n = 5; fmt.Println(a.n); fmt.Println(b.n) }", vec!["2", "5"]),
    pointer_embedded_explicit_inner_access_runtime =>
        ("package main; import \"fmt\"; type inner struct { count int }; type outer struct { *inner }; func main() { o := outer{inner: &inner{count: 4}}; fmt.Println(o.inner.count) }", vec!["4"]),
    value_embedded_zero_value_promoted_field_runtime =>
        ("package main; import \"fmt\"; type inner struct { n int }; type outer struct { inner }; func main() { var o outer; fmt.Println(o.n) }", vec!["0"]),
    nested_pointer_middle_embedding_runtime =>
        ("package main; import \"fmt\"; type inner struct { count int }; type middle struct { *inner }; type outer struct { middle }; func main() { o := outer{middle: middle{inner: &inner{count: 15}}}; fmt.Println(o.count) }", vec!["15"]),
    promoted_method_value_from_outer_runtime =>
        ("package main; import \"fmt\"; type inner struct { n int }; func (i inner) total() int { return i.n }; type outer struct { inner }; func main() { o := outer{inner: inner{n: 6}}; fn := o.total; fmt.Println(fn()) }", vec!["6"]),
    embedded_anonymous_struct_type_runtime =>
        ("package main; import \"fmt\"; type outer struct { struct { x int; y int } }; func main() { o := outer{struct { x int; y int }{x: 1, y: 2}}; fmt.Println(o.x + o.y) }", vec!["3"]),
    two_embedded_same_field_name_requires_qualifier_runtime =>
        ("package main; import \"fmt\"; type left struct { id int }; type right struct { id int }; type pair struct { left; right }; func main() { p := pair{left: left{id: 1}, right: right{id: 2}}; fmt.Println(p.left.id); fmt.Println(p.right.id) }", vec!["1", "2"]),
    embed_pointer_vs_value_mutation_runtime =>
        ("package main; import \"fmt\"; type cell struct { n int }; type wrapValue struct { cell }; type wrapPtr struct { *cell }; func main() { v := wrapValue{cell: cell{n: 1}}; p := wrapPtr{cell: &cell{n: 1}}; v.cell.n = 9; p.n = 8; fmt.Println(v.n); fmt.Println(p.n) }", vec!["9", "8"]),
    deep_two_level_pointer_method_runtime =>
        ("package main; import \"fmt\"; type leaf struct { n int }; func (l *leaf) inc() { l.n++ }; type branch struct { leaf }; type trunk struct { branch }; func main() { t := trunk{branch: branch{leaf: leaf{n: 0}}}; t.inc(); fmt.Println(t.n) }", vec!["1"]),
    embedded_slice_field_len_promoted_runtime =>
        ("package main; import \"fmt\"; type inner struct { items []int }; type outer struct { inner }; func main() { o := outer{inner: inner{items: []int{1, 2, 3}}}; fmt.Println(len(o.items)) }", vec!["3"]),
    embedded_map_field_lookup_promoted_runtime =>
        ("package main; import \"fmt\"; type inner struct { data map[string]int }; type outer struct { inner }; func main() { o := outer{inner: inner{data: map[string]int{\"k\": 4}}}; fmt.Println(o.data[\"k\"]) }", vec!["4"]),
    address_of_promoted_field_mutation_runtime =>
        ("package main; import \"fmt\"; type inner struct { n int }; type outer struct { inner }; func main() { o := outer{inner: inner{n: 1}}; ptr := &o.n; *ptr = 6; fmt.Println(o.n) }", vec!["6"]),
    four_level_two_step_promotion_runtime =>
        ("package main; import \"fmt\"; type d struct { n int }; type c struct { d }; type b struct { c }; type a struct { b }; func main() { v := a{b: b{c: c{d: d{n: 13}}}}; fmt.Println(v.n) }", vec!["13"]),
    embedded_func_field_promoted_call_runtime =>
        ("package main; import \"fmt\"; type inner struct { fn func(int) int }; type outer struct { inner }; func main() { o := outer{inner: inner{fn: func(x int) int { return x * 2 }}}; fmt.Println(o.fn(5)) }", vec!["10"]),
    value_embed_pointer_receiver_chain_runtime =>
        ("package main; import \"fmt\"; type engine struct { rpm int }; func (e *engine) rev() *engine { e.rpm++; return e }; type car struct { engine }; func main() { c := car{engine: engine{rpm: 100}}; c.rev().rev(); fmt.Println(c.rpm) }", vec!["102"]),
    collision_outer_wins_for_same_field_name_runtime =>
        ("package main; import \"fmt\"; type base struct { score int }; type derived struct { base; score int }; func main() { d := derived{base: base{score: 1}, score: 2}; fmt.Println(d.score); fmt.Println(d.base.score) }", vec!["2", "1"]),
    two_level_embedded_interface_field_runtime =>
        ("package main; import \"fmt\"; type speaker interface { talk() string }; type bot struct{}; func (bot) talk() string { return \"beep\" }; type host struct { speaker }; type rack struct { host }; func main() { r := rack{host: host{speaker: bot{}}}; fmt.Println(r.talk()) }", vec!["beep"]),
    embedded_type_name_as_field_selector_runtime =>
        ("package main; import \"fmt\"; type coords struct { x int; y int }; type point struct { coords }; func main() { p := point{coords: coords{x: 3, y: 4}}; fmt.Println(p.coords.x) }", vec!["3"]),
    pointer_embed_nil_safe_explicit_check_runtime =>
        ("package main; import \"fmt\"; type inner struct { n int }; type outer struct { *inner }; func main() { var o outer; fmt.Println(o.inner == nil) }", vec!["true"]),
    nested_value_embed_each_level_runtime =>
        ("package main; import \"fmt\"; type c struct { tag string }; type b struct { c }; type a struct { b }; func main() { v := a{b: b{c: c{tag: \"deep\"}}}; fmt.Println(v.tag) }", vec!["deep"]),
}

go_compile_cases! {
    promoted_field_read_compile =>
        "package main; type inner struct { n int }; type outer struct { inner }; func main() { var o outer; _ = o.n }",
    promoted_field_write_compile =>
        "package main; type inner struct { n int }; type outer struct { inner }; func main() { var o outer; o.n = 1 }",
    promoted_method_call_compile =>
        "package main; type inner struct{}; func (inner) f() {}; type outer struct { inner }; func main() { outer{}.f() }",
    collision_qualifier_left_compile =>
        "package main; type a struct{}; func (a) f() {}; type b struct{}; func (b) f() {}; type c struct { a; b }; func main() { var x c; x.a.f() }",
    collision_qualifier_right_compile =>
        "package main; type a struct{}; func (a) f() {}; type b struct{}; func (b) f() {}; type c struct { a; b }; func main() { var x c; x.b.f() }",
    value_embed_pointer_field_compile =>
        "package main; type inner struct { n int }; type outer struct { *inner }; func main() { o := outer{inner: &inner{n: 1}}; _ = o.n }",
    two_level_promotion_compile =>
        "package main; type leaf struct { v int }; type branch struct { leaf }; type trunk struct { branch }; func main() { _ = trunk{}.v }",
    two_level_method_compile =>
        "package main; type leaf struct{}; func (leaf) f() {}; type branch struct { leaf }; type trunk struct { branch }; func main() { trunk{}.f() }",
    outer_shadow_method_compile =>
        "package main; type inner struct{}; func (inner) m() {}; type outer struct { inner }; func (outer) m() {}; func main() { outer{}.m() }",
    explicit_inner_method_when_shadowed_compile =>
        "package main; type inner struct{}; func (inner) m() {}; type outer struct { inner }; func (outer) m() {}; func main() { var o outer; o.inner.m() }",
    address_of_promoted_field_compile =>
        "package main; type inner struct { n int }; type outer struct { inner }; func main() { var o outer; _ = &o.n }",
    pointer_receiver_on_value_embed_compile =>
        "package main; type inner struct { n int }; func (i *inner) bump() { i.n++ }; type outer struct { inner }; func main() { var o outer; o.bump() }",
    nested_pointer_middle_compile =>
        "package main; type inner struct { n int }; type middle struct { *inner }; type outer struct { middle }; func main() { var o outer; _ = o.n }",
    dual_embedded_fields_compile =>
        "package main; type x struct { a int }; type y struct { b int }; type p struct { x; y }; func main() { var v p; _ = v.a; _ = v.b }",
    embedded_interface_field_compile =>
        "package main; type speaker interface { talk() string }; type host struct { speaker }; func main() { _ = host{} }",
    shadowed_field_both_accessible_compile =>
        "package main; type base struct { id int }; type derived struct { base; id int }; func main() { var d derived; _ = d.id; _ = d.base.id }",
    anonymous_struct_embed_compile =>
        "package main; type outer struct { struct { x int } }; func main() { _ = outer{}.x }",
    four_level_field_promotion_compile =>
        "package main; type d struct { n int }; type c struct { d }; type b struct { c }; type a struct { b }; func main() { _ = a{}.n }",
    embed_value_vs_pointer_type_compile =>
        "package main; type cell struct { n int }; type byValue struct { cell }; type byPtr struct { *cell }; func main() { _ = byValue{}; _ = byPtr{} }",
    deep_two_level_pointer_method_compile =>
        "package main; type leaf struct { n int }; func (l *leaf) inc() {}; type branch struct { leaf }; type trunk struct { branch }; func main() { var t trunk; t.inc() }",
}

macro_rules! go_compile_fail_cases {
    ($($name:ident => $src:expr,)+) => {
        $(#[test] fn $name() { assert!(!compile_ok_check($src)); })+
    };
}

go_compile_fail_cases! {
    ambiguous_promoted_field_compile_fail =>
        "package main; type a struct { x int }; type b struct { x int }; type c struct { a; b }; func main() { var v c; _ = v.x }",
    ambiguous_promoted_method_compile_fail =>
        "package main; type a struct{}; func (a) f() {}; type b struct{}; func (b) f() {}; type c struct { a; b }; func main() { var v c; v.f() }",
}
