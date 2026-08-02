// vybe-test: go/generics_constraints_extended/generic_any_empty_struct
// origin: languages/go/tests/go/test_generics_constraints_extended.rs
// vybe-test-mode: compile

package main
func Empty[T any]() T { var z T
return z }
func main() { _ = Empty[struct{}]() }
