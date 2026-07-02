//! Interface, embedding, and method-set language rules — one rule per test.


go_run_cases! {
    interface_satisfied_by_value => ("package main; import \"fmt\"; type I interface { M() int }; type T int; func (t T) M() int { return int(t) }; func main() { var i I = T(3); fmt.Println(i.M()) }", vec!["3"]),
    interface_satisfied_by_pointer => ("package main; import \"fmt\"; type I interface { M() int }; type T int; func (t *T) M() int { return int(*t) }; func main() { var i I = (*T)(nil); fmt.Println(i == nil) }", vec!["true"]),
    embedding_promotes_method => ("package main; import \"fmt\"; type A struct{}; func (A) Hi() string { return \"hi\" }; type B struct { A }; func main() { var b B; fmt.Println(b.Hi()) }", vec!["hi"]),
    embedding_field_access => ("package main; import \"fmt\"; type Inner struct { N int }; type Outer struct { Inner }; func main() { fmt.Println(Outer{Inner: Inner{N: 4}}.N) }", vec!["4"]),
    pointer_embedding_promotion => ("package main; import \"fmt\"; type A struct{}; func (A) X() int { return 1 }; type B struct { *A }; func main() { b := B{A: &A{}}; fmt.Println(b.X()) }", vec!["1"]),
    method_expression_value => ("package main; import \"fmt\"; type N int; func (n N) Inc() N { return n+1 }; func main() { f := N.Inc; fmt.Println(f(2)) }", vec!["3"]),
    method_expression_pointer => ("package main; import \"fmt\"; type N int; func (n *N) Inc() { *n++ }; func main() { var v N = 1; f := (*N).Inc; f(&v); fmt.Println(v) }", vec!["2"]),
    override_promoted_method => ("package main; import \"fmt\"; type A struct{}; func (A) Name() string { return \"A\" }; type B struct { A }; func (B) Name() string { return \"B\" }; func main() { fmt.Println(B{}.Name()) }", vec!["B"]),
    interface_field_in_struct => ("package main; import \"fmt\"; type I interface { F() }; type S struct { I }; func main() { fmt.Println(S{} .I == nil) }", vec!["true"]),
    type_assertion_to_concrete => ("package main; import \"fmt\"; type I interface { M() }; type T struct{}; func (T) M() {}; func main() { var i I = T{}; fmt.Println(i.(T) == T{}) }", vec!["true"]),
    type_switch_default => ("package main; import \"fmt\"; func main() { switch any(1).(type) { default: fmt.Println(\"d\") } }", vec!["d"]),
    type_switch_case_int => ("package main; import \"fmt\"; func main() { switch v := any(2).(type) { case int: fmt.Println(v); default: fmt.Println(0) } }", vec!["2"]),
    empty_interface_assign_any => ("package main; import \"fmt\"; func main() { var a any = 5; fmt.Println(a) }", vec!["5"]),
    iface_eq_nil_both => ("package main; import \"fmt\"; func main() { var i interface{}; fmt.Println(i == nil) }", vec!["true"]),
}

go_compile_cases! {
    iface_pointer_value_mismatch => "package main; type I interface { M() }; type T int; func (T) M() {}; func main() { var _ I = T(0) }",
    embed_ambiguous_requires_selector => "package main; type A struct{}; func (A) F() {}; type B struct{}; func (B) F() {}; type C struct { A; B }; func main() { var c C; _ = c }",
    struct_anonymous_field_name => "package main; type T struct { int }; func main() { _ = T{int: 1} }",
    interface_method_pointer_receiver_only => "package main; type I interface { M() }; type T int; func (t *T) M() {}; func main() { var _ I = (*T)(nil) }",
    embedding_interface_field => "package main; type I interface { F() }; type S struct { I }; func main() { _ = S{} }",
    method_on_defined_string_type => "package main; type MyString string; func (s MyString) Len() int { return len(s) }; func main() { _ = MyString(\"a\").Len() }",
}
