// vybe-test: go/generics_functions/generic_type_set_union
// origin: languages/go/tests/go/test_generics_functions.rs
// vybe-test-mode: compile

package main
type Number interface { ~int | ~float64 }
func Double[T Number](v T) T { return v + v }
func main() { _ = Double(2) }
