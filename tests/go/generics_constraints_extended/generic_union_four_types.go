// vybe-test: go/generics_constraints_extended/generic_union_four_types
// origin: languages/go/tests/go/test_generics_constraints_extended.rs
// vybe-test-mode: compile

package main
func Pick[T int | int8 | int16 | int32](v T) T { return v }
func main() { _ = Pick(int8(1)) }
