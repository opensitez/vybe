//! Generics: functions and types with distinct constraints and instantiations.

go_run_cases! {
    generic_min_int => ("package main; import \"fmt\"; func Min[T ~int](a, b T) T { if a < b { return a }; return b }; func main() { fmt.Println(Min(3, 7)) }", vec!["3"]),
    generic_max_int => ("package main; import \"fmt\"; func Max[T ~int](a, b T) T { if a > b { return a }; return b }; func main() { fmt.Println(Max(3, 7)) }", vec!["7"]),
    generic_slice_len => ("package main; import \"fmt\"; func Len[T any](s []T) int { return len(s) }; func main() { fmt.Println(Len([]int{1,2,3})) }", vec!["3"]),
    generic_pair_swap => ("package main; import \"fmt\"; func Swap[T any](a, b T) (T, T) { return b, a }; func main() { x, y := Swap(1, 2); fmt.Println(x); fmt.Println(y) }", vec!["2", "1"]),
    generic_stack_push_pop => ("package main; import \"fmt\"; type Stack[T any] struct { items []T }; func (s *Stack[T]) Push(v T) { s.items = append(s.items, v) }; func (s *Stack[T]) Pop() T { n := len(s.items)-1; v := s.items[n]; s.items = s.items[:n]; return v }; func main() { var st Stack[int]; st.Push(5); fmt.Println(st.Pop()) }", vec!["5"]),
    generic_map_keys_len => ("package main; import \"fmt\"; func KeysLen[K comparable, V any](m map[K]V) int { return len(m) }; func main() { fmt.Println(KeysLen(map[string]int{\"a\":1})) }", vec!["1"]),
}

go_compile_cases! {
    generic_comparable_map_key => "package main; func Keys[K comparable, V any](m map[K]V) []K { keys := make([]K, 0, len(m)); for k := range m { keys = append(keys, k) }; return keys }; func main() { _ = Keys(map[int]string{1:\"a\"}) }",
    generic_ordered_constraint => "package main; import \"cmp\"; func Clamp[T cmp.Ordered](v, lo, hi T) T { if v < lo { return lo }; if v > hi { return hi }; return v }; func main() { _ = Clamp(3, 1, 5) }",
    generic_type_set_union => "package main; type Number interface { ~int | ~float64 }; func Double[T Number](v T) T { return v + v }; func main() { _ = Double(2) }",
    generic_pointer_constraint => "package main; func Zero[T any](p *T) { var z T; *p = z }; func main() { var x int; Zero(&x) }",
    generic_method_on_type => "package main; type Box[T any] struct { V T }; func (b Box[T]) Get() T { return b.V }; func main() { _ = Box[int]{V:1}.Get() }",
}
