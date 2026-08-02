// vybe-test: go/lang_generics_semantics/constraint_union_three_types
// origin: languages/go/tests/go/test_lang_generics_semantics.rs
// vybe-test-mode: compile

package main
func F[T int | float64 | string](v T) T { return v }
func main() { _ = F(1) }
