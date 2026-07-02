//! Generics: struct types, interfaces, receiver methods, and constraint interfaces.
//!
//! Complements `test_generics_functions.rs` (generic functions and function constraints).


go_run_cases! {
    // --- generic struct types ---
    generic_pair_homogeneous_fields => (
        "package main; import \"fmt\"; type Pair[T any] struct { First, Second T }; func main() { p := Pair[int]{First: 3, Second: 7}; fmt.Println(p.First); fmt.Println(p.Second) }",
        vec!["3", "7"]
    ),
    generic_pair_heterogeneous_type_params => (
        "package main; import \"fmt\"; type Pair[A, B any] struct { First A; Second B }; func main() { p := Pair[int, string]{First: 9, Second: \"go\"}; fmt.Println(p.First); fmt.Println(p.Second) }",
        vec!["9", "go"]
    ),
    generic_triple_three_type_parameters => (
        "package main; import \"fmt\"; type Triple[A, B, C any] struct { A A; B B; C C }; func main() { t := Triple[int, string, bool]{A: 1, B: \"x\", C: true}; fmt.Println(t.A); fmt.Println(t.B); fmt.Println(t.C) }",
        vec!["1", "x", "true"]
    ),
    generic_node_linked_list_fields => (
        "package main; import \"fmt\"; type Node[T any] struct { Val T; Next *Node[T] }; func main() { head := &Node[int]{Val: 1, Next: &Node[int]{Val: 2}}; fmt.Println(head.Val); fmt.Println(head.Next.Val) }",
        vec!["1", "2"]
    ),
    generic_queue_enqueue_dequeue => (
        "package main; import \"fmt\"; type Queue[T any] struct { items []T }; func (q *Queue[T]) Enqueue(v T) { q.items = append(q.items, v) }; func (q *Queue[T]) Dequeue() T { v := q.items[0]; q.items = q.items[1:]; return v }; func main() { var q Queue[int]; q.Enqueue(10); q.Enqueue(20); fmt.Println(q.Dequeue()); fmt.Println(q.Dequeue()) }",
        vec!["10", "20"]
    ),
    generic_set_comparable_membership => (
        "package main; import \"fmt\"; type Set[T comparable] struct { m map[T]struct{} }; func (s *Set[T]) Add(v T) { if s.m == nil { s.m = make(map[T]struct{}) }; s.m[v] = struct{}{} }; func (s *Set[T]) Has(v T) bool { _, ok := s.m[v]; return ok }; func main() { var s Set[int]; s.Add(3); fmt.Println(s.Has(3)); fmt.Println(s.Has(4)) }",
        vec!["true", "false"]
    ),
    generic_option_some_present => (
        "package main; import \"fmt\"; type Option[T any] struct { Value T; Present bool }; func Some[T any](v T) Option[T] { return Option[T]{Value: v, Present: true} }; func main() { o := Some(42); fmt.Println(o.Present); fmt.Println(o.Value) }",
        vec!["true", "42"]
    ),
    generic_wrapper_single_field => (
        "package main; import \"fmt\"; type Wrapper[T any] struct { V T }; func main() { w := Wrapper[string]{V: \"vybe\"}; fmt.Println(w.V) }",
        vec!["vybe"]
    ),
    generic_matrix_nested_slice => (
        "package main; import \"fmt\"; type Matrix[T any] struct { Rows [][]T }; func main() { m := Matrix[int]{Rows: [][]int{{1, 2}, {3, 4}}}; fmt.Println(len(m.Rows)); fmt.Println(m.Rows[1][0]) }",
        vec!["2", "3"]
    ),
    generic_chan_wrap_send_recv => (
        "package main; import \"fmt\"; type ChanWrap[T any] struct { Ch chan T }; func main() { w := ChanWrap[int]{Ch: make(chan int, 1)}; w.Ch <- 5; fmt.Println(<-w.Ch) }",
        vec!["5"]
    ),

    // --- generic interfaces ---
    generic_container_interface_impl => (
        "package main; import \"fmt\"; type Container[T any] interface { Add(T); Size() int }; type SliceBox[T any] struct { items []T }; func (s *SliceBox[T]) Add(v T) { s.items = append(s.items, v) }; func (s *SliceBox[T]) Size() int { return len(s.items) }; func main() { var c Container[int] = &SliceBox[int]{}; c.Add(5); c.Add(8); fmt.Println(c.Size()) }",
        vec!["2"]
    ),
    generic_reader_writer_embedded_iface => (
        "package main; import \"fmt\"; type Reader[T any] interface { Read() T }; type Writer[T any] interface { Write(T) }; type ReadWriter[T any] interface { Reader[T]; Writer[T] }; type Buffer[T any] struct { data []T }; func (b *Buffer[T]) Read() T { v := b.data[0]; b.data = b.data[1:]; return v }; func (b *Buffer[T]) Write(v T) { b.data = append(b.data, v) }; func main() { var rw ReadWriter[int] = &Buffer[int]{}; rw.Write(3); fmt.Println(rw.Read()) }",
        vec!["3"]
    ),
    generic_comparer_ordered_constraint => (
        "package main; import \"fmt\"; import \"cmp\"; type Comparer[T cmp.Ordered] interface { Less(a, b T) bool }; type IntCmp struct{}; func (IntCmp) Less(a, b int) bool { return a < b }; func main() { var c Comparer[int] = IntCmp{}; fmt.Println(c.Less(2, 5)) }",
        vec!["true"]
    ),

    // --- type parameters on receiver methods ---
    generic_holder_value_receiver_peek => (
        "package main; import \"fmt\"; type Holder[T any] struct { items []T }; func (h Holder[T]) Peek() T { return h.items[0] }; func main() { fmt.Println(Holder[int]{items: []int{11}}.Peek()) }",
        vec!["11"]
    ),
    generic_vault_pointer_receiver_store => (
        "package main; import \"fmt\"; type Vault[T any] struct { V T }; func (v *Vault[T]) Store(x T) { v.V = x }; func main() { vault := Vault[int]{V: 1}; vault.Store(99); fmt.Println(vault.V) }",
        vec!["99"]
    ),
    generic_pair_receiver_swap_fields => (
        "package main; import \"fmt\"; type Pair[T any] struct { First, Second T }; func (p *Pair[T]) Swap() { p.First, p.Second = p.Second, p.First }; func main() { p := Pair[int]{First: 1, Second: 2}; p.Swap(); fmt.Println(p.First); fmt.Println(p.Second) }",
        vec!["2", "1"]
    ),
    generic_list_append_method => (
        "package main; import \"fmt\"; type List[T any] struct { items []T }; func (l *List[T]) Append(v T) { l.items = append(l.items, v) }; func (l List[T]) At(i int) T { return l.items[i] }; func main() { var l List[string]; l.Append(\"a\"); l.Append(\"b\"); fmt.Println(l.At(1)) }",
        vec!["b"]
    ),
    generic_counter_increment_method => (
        "package main; import \"fmt\"; type Counter[T ~int] struct { n T }; func (c *Counter[T]) Inc() { c.n++ }; func (c Counter[T]) Value() T { return c.n }; func main() { c := Counter[int]{n: 4}; c.Inc(); fmt.Println(c.Value()) }",
        vec!["5"]
    ),
    generic_calculator_twice_method => (
        "package main; import \"fmt\"; type Numeric interface { ~int | ~float64 }; type Calculator[T Numeric] struct{}; func (Calculator[T]) Twice(v T) T { return v + v }; func main() { fmt.Println(Calculator[int]{}.Twice(6)) }",
        vec!["12"]
    ),

    // --- constraint interfaces ---
    generic_signed_tilde_constraint => (
        "package main; import \"fmt\"; type Signed interface { ~int | ~int64 }; type Abs[T Signed] struct{}; func (Abs[T]) Negate(v T) T { return -v }; func main() { fmt.Println(Abs[int]{}.Negate(7)) }",
        vec!["-7"]
    ),
    generic_cache_comparable_key => (
        "package main; import \"fmt\"; type Cache[K comparable, V any] struct { m map[K]V }; func (c *Cache[K, V]) Put(k K, v V) { if c.m == nil { c.m = make(map[K]V) }; c.m[k] = v }; func (c Cache[K, V]) Get(k K) (V, bool) { v, ok := c.m[k]; return v, ok }; func main() { var c Cache[string, int]; c.Put(\"x\", 3); v, ok := c.Get(\"x\"); fmt.Println(ok); fmt.Println(v) }",
        vec!["true", "3"]
    ),
    generic_stringer_constraint_method => (
        "package main; import \"fmt\"; type Stringer interface { String() string }; type Show[T Stringer] struct{}; func (Show[T]) Display(v T) string { return v.String() }; type Tag struct { Label string }; func (t Tag) String() string { return t.Label }; func main() { fmt.Println(Show[Tag]{}.Display(Tag{Label: \"ok\"})) }",
        vec!["ok"]
    ),
    generic_union_constraint_float => (
        "package main; import \"fmt\"; type Real interface { ~int | ~float64 }; type Half[T Real] struct{}; func (Half[T]) Of(v T) T { return v / 2 }; func main() { fmt.Println(Half[float64]{}.Of(5.0)) }",
        vec!["2.5"]
    ),
    generic_embed_ordered_in_constraint => (
        "package main; import \"fmt\"; import \"cmp\"; type Ordered = cmp.Ordered; type Sorter[T Ordered] struct{}; func (Sorter[T]) IsLess(a, b T) bool { return a < b }; func main() { fmt.Println(Sorter[int]{}.IsLess(1, 3)) }",
        vec!["true"]
    ),
}

go_compile_cases! {
    generic_struct_literal_zero_value => "package main; type Cell[T any] struct { V T }; func main() { var c Cell[int]; _ = c }",
    generic_interface_type_param_decl => "package main; type Storer[T any] interface { Store(T) }; func main() {}",
    generic_method_chained_on_pointer => "package main; type Node[T any] struct { Val T; Next *Node[T] }; func (n *Node[T]) Link(next *Node[T]) *Node[T] { n.Next = next; return n }; func main() { _ = (&Node[int]{Val: 1}).Link(&Node[int]{Val: 2}) }",
    generic_nested_type_parameter => "package main; type Outer[T any] struct { Inner struct { V T } }; func main() { _ = Outer[int]{Inner: struct{ V int }{V: 1}} }",
    generic_alias_with_type_param => "package main; type List[T any] []T; func main() { var xs List[int]; _ = xs }",
    generic_struct_constraint_on_field => "package main; type Keyer interface { Key() string }; type Named struct { Name string }; func (n Named) Key() string { return n.Name }; type Entry[T Keyer] struct { Item T }; func main() { _ = Entry[Named]{Item: Named{Name: \"a\"}} }",
    generic_multi_constraint_interface_embed => "package main; import \"cmp\"; type KeyedOrdered[T cmp.Ordered] interface { cmp.Ordered; Key() T }; func main() {}",
}
