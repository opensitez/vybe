// vybe-test: go/generics_types/generic_nested_type_parameter
// origin: languages/go/tests/go/test_generics_types.rs
// vybe-test-mode: compile

package main
type Outer[T any] struct { Inner struct { V T } }
func main() { _ = Outer[int]{Inner: struct{ V int }{V: 1}} }
