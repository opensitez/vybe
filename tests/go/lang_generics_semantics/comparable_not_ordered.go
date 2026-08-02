// vybe-test: go/lang_generics_semantics/comparable_not_ordered
// origin: languages/go/tests/go/test_lang_generics_semantics.rs
// vybe-test-mode: compile

package main
func Eq[T comparable](a, b T) bool { return a == b }
func main() { _ = Eq([1]int{1}, [1]int{1}) }
