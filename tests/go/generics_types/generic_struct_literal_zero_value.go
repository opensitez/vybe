// vybe-test: go/generics_types/generic_struct_literal_zero_value
// origin: languages/go/tests/go/test_generics_types.rs
// vybe-test-mode: compile

package main
type Cell[T any] struct { V T }
func main() { var c Cell[int]
_ = c }
