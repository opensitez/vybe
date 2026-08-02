// vybe-test: go/lang_generics_semantics/generic_slice_append
// origin: languages/go/tests/go/test_lang_generics_semantics.rs
// vybe-test-mode: compile

package main
func Append[T any](s []T, v T) []T { return append(s, v) }
func main() { _ = Append([]int{1}, 2) }
