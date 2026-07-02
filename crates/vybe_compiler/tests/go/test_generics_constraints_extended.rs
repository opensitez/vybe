//! Generics constraints extended: comparable map keys, cmp.Ordered sorting,
//! any vs comparable, union type sets, tilde (~) constraints on structs, and
//! type parameters on methods — distinct from `test_generics_functions.rs`,
//! `test_generics_types.rs`, and `test_lang_generics_semantics.rs`.


go_run_cases! {
    generic_comparable_map_lookup_string => (
        "package main; import \"fmt\"; func Get[K comparable, V any](m map[K]V, k K) (V, bool) { v, ok := m[k]; return v, ok }; func main() { v, ok := Get(map[string]int{\"go\": 1}, \"go\"); fmt.Println(v); fmt.Println(ok) }",
        vec!["1", "true"]
    ),
    generic_comparable_map_lookup_missing => (
        "package main; import \"fmt\"; func Get[K comparable, V any](m map[K]V, k K) (V, bool) { v, ok := m[k]; return v, ok }; func main() { _, ok := Get(map[int]string{1: \"a\"}, 9); fmt.Println(ok) }",
        vec!["false"]
    ),
    generic_comparable_map_keys_to_slice => (
        "package main; import \"fmt\"; func KeyList[K comparable, V any](m map[K]V) []K { keys := make([]K, 0, len(m)); for k := range m { keys = append(keys, k) }; return keys }; func main() { fmt.Println(len(KeyList(map[int]bool{1: true, 2: false}))) }",
        vec!["2"]
    ),
    generic_comparable_struct_key_map => (
        "package main; import \"fmt\"; type Key struct { ID int }; func Has[K comparable, V any](m map[K]V, k K) bool { _, ok := m[k]; return ok }; func main() { k := Key{ID: 1}; fmt.Println(Has(map[Key]string{k: \"v\"}, k)) }",
        vec!["true"]
    ),
    generic_comparable_array_key => (
        "package main; import \"fmt\"; func LenMap[K comparable, V any](m map[K]V) int { return len(m) }; func main() { fmt.Println(LenMap(map[[2]int]string{[2]int{1, 2}: \"pair\"})) }",
        vec!["1"]
    ),
    generic_ordered_sort_ascending => (
        "package main; import \"fmt\"; import \"cmp\"; import \"slices\"; func SortAsc[T cmp.Ordered](s []T) { slices.Sort(s) }; func main() { data := []int{3, 1, 2}; SortAsc(data); fmt.Println(data[0]); fmt.Println(data[2]) }",
        vec!["1", "3"]
    ),
    generic_ordered_max_float64 => (
        "package main; import \"fmt\"; import \"cmp\"; func Max[T cmp.Ordered](a, b T) T { if cmp.Less(a, b) { return b }; return a }; func main() { fmt.Println(Max(1.5, 2.5)) }",
        vec!["2.5"]
    ),
    generic_ordered_min_string => (
        "package main; import \"fmt\"; import \"cmp\"; func Min[T cmp.Ordered](a, b T) T { if cmp.Less(a, b) { return a }; return b }; func main() { fmt.Println(Min(\"zebra\", \"apple\")) }",
        vec!["apple"]
    ),
    generic_ordered_clamp_int => (
        "package main; import \"fmt\"; import \"cmp\"; func Clamp[T cmp.Ordered](v, lo, hi T) T { if cmp.Less(v, lo) { return lo }; if cmp.Less(hi, v) { return hi }; return v }; func main() { fmt.Println(Clamp(99, 0, 10)) }",
        vec!["10"]
    ),
    generic_ordered_sort_strings => (
        "package main; import \"fmt\"; import \"slices\"; func SortStrings[T ~string](s []T) { slices.Sort(s) }; func main() { names := []string{\"go\", \"vybe\", \"lang\"}; SortStrings(names); fmt.Println(names[0]) }",
        vec!["go"]
    ),
    generic_any_identity_bool => (
        "package main; import \"fmt\"; func ID[T any](v T) T { return v }; func main() { fmt.Println(ID(true)) }",
        vec!["true"]
    ),
    generic_any_slice_reverse => (
        "package main; import \"fmt\"; func Reverse[T any](s []T) { for i, j := 0, len(s)-1; i < j; i, j = i+1, j-1 { s[i], s[j] = s[j], s[i] } }; func main() { a := []int{1, 2, 3}; Reverse(a); fmt.Println(a[0]); fmt.Println(a[2]) }",
        vec!["3", "1"]
    ),
    generic_comparable_equal_slices => (
        "package main; import \"fmt\"; func Eq[T comparable](a, b T) bool { return a == b }; func main() { fmt.Println(Eq(1, 1)); fmt.Println(Eq(1, 2)) }",
        vec!["true", "false"]
    ),
    generic_comparable_not_for_ordered_sort => (
        "package main; import \"fmt\"; func CountMap[K comparable, V any](m map[K]V) int { return len(m) }; func main() { fmt.Println(CountMap(map[bool]int{true: 1, false: 2})) }",
        vec!["2"]
    ),
    generic_union_int_branch => (
        "package main; import \"fmt\"; func Twice[T int | int64](v T) T { return v + v }; func main() { fmt.Println(Twice(5)) }",
        vec!["10"]
    ),
    generic_union_string_branch => (
        "package main; import \"fmt\"; func Len[T string | []byte](v T) int { return len(v) }; func main() { fmt.Println(Len(\"go\")) }",
        vec!["2"]
    ),
    generic_union_float64 => (
        "package main; import \"fmt\"; func Double[T float32 | float64](v T) T { return v * 2 }; func main() { fmt.Println(Double(2.5)) }",
        vec!["5"]
    ),
    generic_union_three_numeric => (
        "package main; import \"fmt\"; func Sign[T int | int64 | float64](v T) int { if v < 0 { return -1 }; if v > 0 { return 1 }; return 0 }; func main() { fmt.Println(Sign(int64(-3))) }",
        vec!["-1"]
    ),
    generic_tilde_myint_constraint => (
        "package main; import \"fmt\"; type MyInt int; func AddOne[T ~int](v T) T { return v + 1 }; func main() { fmt.Println(AddOne(MyInt(4))) }",
        vec!["5"]
    ),
    generic_tilde_mystring_upper => (
        "package main; import \"fmt\"; import \"strings\"; type Label string; func Upper[T ~string](s T) T { return T(strings.ToUpper(string(s))) }; func main() { fmt.Println(Upper(Label(\"go\"))) }",
        vec!["GO"]
    ),
    generic_tilde_struct_slice_len => (
        "package main; import \"fmt\"; type Ints []int; func Total[S ~[]int](s S) int { sum := 0; for _, v := range s { sum += v }; return sum }; func main() { fmt.Println(Total(Ints{1, 2, 3})) }",
        vec!["6"]
    ),
    generic_tilde_struct_map_len => (
        "package main; import \"fmt\"; type StrMap map[string]int; func Size[M ~map[string]int](m M) int { return len(m) }; func main() { fmt.Println(Size(StrMap{\"a\": 1, \"b\": 2})) }",
        vec!["2"]
    ),
    generic_method_value_receiver_get => (
        "package main; import \"fmt\"; type Cell[T any] struct { V T }; func (c Cell[T]) Get() T { return c.V }; func main() { fmt.Println(Cell[int]{V: 42}.Get()) }",
        vec!["42"]
    ),
    generic_method_pointer_receiver_set => (
        "package main; import \"fmt\"; type Cell[T any] struct { V T }; func (c *Cell[T]) Set(v T) { c.V = v }; func main() { c := Cell[string]{}; c.Set(\"ok\"); fmt.Println(c.V) }",
        vec!["ok"]
    ),
    generic_method_comparable_contains => (
        "package main; import \"fmt\"; type Set[T comparable] struct { items []T }; func (s *Set[T]) Add(v T) { s.items = append(s.items, v) }; func (s Set[T]) Has(v T) bool { for _, x := range s.items { if x == v { return true } }; return false }; func main() { var st Set[int]; st.Add(7); fmt.Println(st.Has(7)); fmt.Println(st.Has(8)) }",
        vec!["true", "false"]
    ),
    generic_method_ordered_max_in_slice => (
        "package main; import \"fmt\"; import \"cmp\"; type Stats[T cmp.Ordered] struct { data []T }; func (s Stats[T]) Max() T { m := s.data[0]; for _, v := range s.data[1:] { if cmp.Less(m, v) { m = v } }; return m }; func main() { fmt.Println(Stats[int]{data: []int{1, 9, 3}}.Max()) }",
        vec!["9"]
    ),
    generic_method_any_len => (
        "package main; import \"fmt\"; type Bag[T any] struct { items []T }; func (b Bag[T]) Len() int { return len(b.items) }; func main() { fmt.Println(Bag[string]{items: []string{\"a\", \"b\"}}.Len()) }",
        vec!["2"]
    ),
    generic_method_chained_on_type_param => (
        "package main; import \"fmt\"; type Num[T ~int] struct { V T }; func (n Num[T]) Inc() Num[T] { n.V++; return n }; func (n Num[T]) Value() T { return n.V }; func main() { fmt.Println(Num[int]{V: 1}.Inc().Value()) }",
        vec!["2"]
    ),
    generic_map_comparable_key_insert => (
        "package main; import \"fmt\"; func Put[K comparable, V any](m map[K]V, k K, v V) { m[k] = v }; func main() { m := map[rune]int{}; Put(m, 'x', 1); fmt.Println(m['x']) }",
        vec!["1"]
    ),
    generic_ordered_binary_search => (
        "package main; import \"fmt\"; import \"cmp\"; import \"slices\"; func Find[T cmp.Ordered](s []T, target T) (int, bool) { return slices.BinarySearch(s, target) }; func main() { i, ok := Find([]int{1, 3, 5}, 3); fmt.Println(i); fmt.Println(ok) }",
        vec!["1", "true"]
    ),
    generic_union_print_type_int => (
        "package main; import \"fmt\"; func Tag[T int | string](v T) string { switch any(v).(type) { case int: return \"int\"; default: return \"string\" } }; func main() { fmt.Println(Tag(1)) }",
        vec!["int"]
    ),
    generic_union_print_type_string => (
        "package main; import \"fmt\"; func Tag[T int | string](v T) string { switch any(v).(type) { case int: return \"int\"; default: return \"string\" } }; func main() { fmt.Println(Tag(\"x\")) }",
        vec!["string"]
    ),
    generic_tilde_byte_slice_sum => (
        "package main; import \"fmt\"; type Bytes []byte; func SumBytes[B ~[]byte](b B) int { s := 0; for _, c := range b { s += int(c) }; return s }; func main() { fmt.Println(SumBytes(Bytes{'a', 'b'})) }",
        vec!["195"]
    ),
    generic_comparable_interface_key => (
        "package main; import \"fmt\"; func Count[K comparable, V any](m map[K]V) int { return len(m) }; func main() { type I interface { ~int }; fmt.Println(Count(map[int]string{1: \"a\"})) }",
        vec!["1"]
    ),
    generic_any_nil_pointer_zero => (
        "package main; import \"fmt\"; func ZeroPtr[T any]() *T { return nil }; func main() { fmt.Println(ZeroPtr[int]() == nil) }",
        vec!["true"]
    ),
    generic_method_union_constraint => (
        "package main; import \"fmt\"; type Converter[T int | float64] struct { Factor T }; func (c Converter[T]) Scale(v T) T { return v * c.Factor }; func main() { fmt.Println(Converter[int]{Factor: 3}.Scale(4)) }",
        vec!["12"]
    ),
    generic_ordered_three_way_compare => (
        "package main; import \"fmt\"; import \"cmp\"; func Compare3[T cmp.Ordered](a, b, c T) T { if cmp.Less(a, b) { return a }; if cmp.Less(b, c) { return b }; return c }; func main() { fmt.Println(Compare3(5, 2, 8)) }",
        vec!["2"]
    ),
    generic_comparable_map_equal_keys => (
        "package main; import \"fmt\"; func SameKeys[K comparable, V any](a, b map[K]V) bool { if len(a) != len(b) { return false }; for k := range a { if _, ok := b[k]; !ok { return false } }; return true }; func main() { a := map[string]int{\"x\": 1}; b := map[string]int{\"x\": 2}; fmt.Println(SameKeys(a, b)) }",
        vec!["true"]
    ),
    generic_tilde_custom_stringer => (
        "package main; import \"fmt\"; type MyString string; func Quote[T ~string](s T) string { return \"\\\"\" + string(s) + \"\\\"\" }; func main() { fmt.Println(Quote(MyString(\"go\"))) }",
        vec!["\"go\""]
    ),
    generic_method_on_generic_interface => (
        "package main; import \"fmt\"; type Stringer[T any] interface { Format() string }; type Item[T any] struct { V T }; func (i Item[T]) Format() string { return fmt.Sprintf(\"%v\", i.V) }; func Print[T any](s Stringer[T]) { fmt.Println(s.Format()) }; func main() { Print(Item[int]{V: 7}) }",
        vec!["7"]
    ),
    generic_any_append_to_slice => (
        "package main; import \"fmt\"; func Append[T any](s []T, vals ...T) []T { return append(s, vals...) }; func main() { fmt.Println(len(Append([]int{1}, 2, 3))) }",
        vec!["3"]
    ),
    generic_comparable_delete_key => (
        "package main; import \"fmt\"; func Del[K comparable, V any](m map[K]V, k K) { delete(m, k) }; func main() { m := map[int]string{1: \"a\", 2: \"b\"}; Del(m, 1); fmt.Println(len(m)) }",
        vec!["1"]
    ),
}

go_compile_cases! {
    generic_comparable_chan_key => "package main; func KeyType[K comparable]() K { var z K; return z }; func main() { _ = KeyType[chan int]() }",
    generic_ordered_uint_constraint => "package main; import \"cmp\"; func MinU[T cmp.Ordered](a, b T) T { if a < b { return a }; return b }; func main() { _ = MinU(uint(1), uint(2)) }",
    generic_union_four_types => "package main; func Pick[T int | int8 | int16 | int32](v T) T { return v }; func main() { _ = Pick(int8(1)) }",
    generic_tilde_struct_field_type => "package main; type Score int; func Double[T ~int](v T) T { return v * 2 }; func main() { _ = Double(Score(3)) }",
    generic_method_pointer_on_generic_struct => "package main; type Node[T any] struct { Val T; Next *Node[T] }; func (n *Node[T]) Link(next *Node[T]) { n.Next = next }; func main() { a, b := &Node[int]{}, &Node[int]{}; a.Link(b) }",
    generic_any_func_param => "package main; func Apply[T any](f func(T) T, v T) T { return f(v) }; func main() { _ = Apply(func(x int) int { return x + 1 }, 1) }",
    generic_comparable_slice_index => "package main; func Index[T comparable](s []T, v T) int { for i, x := range s { if x == v { return i } }; return -1 }; func main() { _ = Index([]string{\"a\"}, \"a\") }",
    generic_ordered_heap_not_needed_sortfunc => "package main; import \"cmp\"; import \"slices\"; func SortDesc[T cmp.Ordered](s []T) { slices.SortFunc(s, func(a, b T) int { if a > b { return -1 }; if a < b { return 1 }; return 0 }) }; func main() { data := []int{1, 3, 2}; SortDesc(data) }",
    generic_union_pointer_types => "package main; func Deref[T *int | *string](p T) interface{} { switch v := any(p).(type) { case *int: return *v; default: return *v.(*string) } }; func main() { x := 1; _ = Deref(&x) }",
    generic_tilde_map_with_custom_type => "package main; type M map[string]int; func KeysLen[T ~map[string]int](m T) int { return len(m) }; func main() { _ = KeysLen(M{\"a\": 1}) }",
    generic_method_value_on_comparable_set => "package main; type IDSet[T comparable] map[T]struct{}; func (s IDSet[T]) Add(v T) { s[v] = struct{}{} }; func main() { m := IDSet[int]{}; m.Add(1) }",
    generic_nested_type_parameter => "package main; type Outer[T any] struct { Inner[T] }; type Inner[U any] struct { V U }; func main() { _ = Outer[int]{Inner: Inner[int]{V: 1}} }",
    generic_comparable_array_element => "package main; func HasArray[T comparable](s []T, v T) bool { for _, x := range s { if x == v { return true } }; return false }; func main() { _ = HasArray([][1]int{{1}}, [1]int{1}) }",
    generic_any_type_switch => "package main; func Kind[T any](v T) string { switch any(v).(type) { case int: return \"int\"; case string: return \"string\"; default: return \"other\" } }; func main() { _ = Kind(1.0) }",
    generic_ordered_string_compare => "package main; import \"cmp\"; func Sorted[T cmp.Ordered](a, b T) bool { return cmp.Less(a, b) }; func main() { _ = Sorted(\"a\", \"b\") }",
    generic_union_signed_unsigned => "package main; func PickNum[T int | uint](v T) T { return v }; func main() { _ = PickNum(-3); _ = PickNum(uint(3)) }",
    generic_tilde_float_type => "package main; type Real float64; func Square[T ~float64](v T) T { return v * v }; func main() { _ = Square(Real(2)) }",
    generic_method_multiple_type_params => "package main; type Pair[A, B any] struct { A A; B B }; func (p Pair[A, B]) Swap() Pair[B, A] { return Pair[B, A]{A: p.B, B: p.A} }; func main() { _ = Pair[int, string]{1, \"x\"}.Swap() }",
    generic_comparable_map_range => "package main; func SumValues[K comparable, V ~int](m map[K]V) int { s := 0; for _, v := range m { s += int(v) }; return s }; func main() { _ = SumValues(map[string]int8{\"a\": 1}) }",
    generic_any_empty_struct => "package main; func Empty[T any]() T { var z T; return z }; func main() { _ = Empty[struct{}]() }",
    generic_interface_constraint_two_methods => "package main; type RW interface { Read() int; Write(int) }; func Use[T RW](v T) { v.Write(1); _ = v.Read() }; type S struct { n int }; func (s *S) Read() int { return s.n }; func (s *S) Write(n int) { s.n = n }; func main() { var x S; Use(&x) }",
    generic_method_on_constraint_interface => "package main; import \"cmp\"; type Sorter[T cmp.Ordered] interface { Sort() }; type Ints []int; func (s Ints) Sort() { for i := 0; i < len(s); i++ { for j := i+1; j < len(s); j++ { if s[j] < s[i] { s[i], s[j] = s[j], s[i] } } } }; func Run[T cmp.Ordered, S Sorter[T]](s S) { s.Sort() }; func main() { data := Ints{3, 1, 2}; Run(data) }",
    generic_tilde_slice_to_custom => "package main; type Names []string; func First[N ~[]string](n N) string { if len(n) == 0 { return \"\" }; return n[0] }; func main() { _ = First(Names{\"go\"}) }",
    generic_comparable_func_not_allowed_compile => "package main; func Bad[K comparable]() {}; func main() { type F func(); _ = F }",
}
