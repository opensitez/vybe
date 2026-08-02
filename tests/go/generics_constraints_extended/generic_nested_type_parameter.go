// vybe-test: go/generics_constraints_extended/generic_nested_type_parameter
// origin: languages/go/tests/go/test_generics_constraints_extended.rs
// vybe-test-mode: compile

package main
type Outer[T any] struct { Inner[T] }
type Inner[U any] struct { V U }
func main() { _ = Outer[int]{Inner: Inner[int]{V: 1}} }
