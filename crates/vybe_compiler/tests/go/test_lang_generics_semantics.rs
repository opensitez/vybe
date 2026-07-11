//! Generics language semantics — one distinct constraint/syntax behavior per test.

go_run_cases! {
    generic_func_identity_int => ("package main; import \"fmt\"; func ID[T any](v T) T { return v }; func main() { fmt.Println(ID(7)) }", vec!["7"]),
    generic_func_two_type_params => ("package main; import \"fmt\"; func Pair[A any, B any](a A, b B) (A, B) { return a, b }; func main() { x, y := Pair(1, \"x\"); fmt.Println(x); fmt.Println(y) }", vec!["1", "x"]),
    generic_struct_field => ("package main; import \"fmt\"; type Box[T any] struct { V T }; func main() { fmt.Println(Box[int]{V: 2}.V) }", vec!["2"]),
    generic_method_on_type => ("package main; import \"fmt\"; type S[T comparable] struct { V T }; func (s S[T]) Same(u T) bool { return s.V == u }; func main() { fmt.Println(S[int]{1}.Same(1)) }", vec!["true"]),
    comparable_map_keys_generic => ("package main; import \"fmt\"; func Keys[M ~map[K]V, K comparable, V any](m M) int { return len(m) }; func main() { fmt.Println(Keys(map[string]int{\"a\":1})) }", vec!["1"]),
    ordered_constraint_min => ("package main; import \"fmt\"; import \"cmp\"; func Smallest[T cmp.Ordered](a, b T) T { if cmp.Less(a,b) { return a }; return b }; func main() { fmt.Println(Smallest(3,9)) }", vec!["3"]),
    tilde_constraint_slice_len => ("package main; import \"fmt\"; func Len[S ~[]E, E any](s S) int { return len(s) }; func main() { fmt.Println(Len([]int{1,2,3})) }", vec!["3"]),
    interface_constraint_method => ("package main; import \"fmt\"; type Stringer interface { String() string }; func Print[T Stringer](v T) { fmt.Println(v.String()) }; type My int; func (m My) String() string { return \"m\" }; func main() { Print(My(0)) }", vec!["m"]),
    union_constraint_type_set => ("package main; import \"fmt\"; func Describe[T int | string](v T) { fmt.Printf(\"%v\", v) }; func main() { Describe(1) }", vec!["1"]),
    generic_instantiation_explicit => ("package main; import \"fmt\"; func Zero[T any]() T { var z T; return z }; func main() { fmt.Println(Zero[int]()) }", vec!["0"]),
}

go_compile_cases! {
    generic_nested_type_param => "package main; type Outer[T any] struct { Inner[T] }; type Inner[U any] struct { V U }; func main() { _ = Outer[int]{} }",
    generic_pointer_receiver => "package main; type P[T any] struct { v T }; func (p *P[T]) Set(v T) { p.v = v }; func main() { p := &P[int]{}; p.Set(1) }",
    generic_interface_embedding => "package main; type I[T any] interface { Get() T }; type J[T any] interface { I[T]; Set(T) }; func main() { var _ J[int] }",
    constraint_union_three_types => "package main; func F[T int | float64 | string](v T) T { return v }; func main() { _ = F(1) }",
    comparable_not_ordered => "package main; func Eq[T comparable](a, b T) bool { return a == b }; func main() { _ = Eq([1]int{1}, [1]int{1}) }",
    generic_slice_append => "package main; func Append[T any](s []T, v T) []T { return append(s, v) }; func main() { _ = Append([]int{1}, 2) }",
    generic_map_make => "package main; func MakeMap[K comparable, V any]() map[K]V { return make(map[K]V) }; func main() { _ = MakeMap[string, int]() }",
    type_inference_from_args => "package main; func Dup[T any](v T) (T,T) { return v, v }; func main() { _, _ = Dup(\"a\") }",
    generic_chan_send_recv => "package main; func Pump[T any](ch chan T, v T) { ch <- v }; func main() { ch := make(chan int,1); Pump(ch, 1) }",
    generic_func_type_param_in_sig => "package main; func Apply[T any, R any](f func(T) R, v T) R { return f(v) }; func main() { _ = Apply(func(int) string { return \"\" }, 1) }",
}
