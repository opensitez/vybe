// vybe-test: go/generics_functions/generic_method_on_type
// origin: languages/go/tests/go/test_generics_functions.rs
// vybe-test-mode: compile

package main
type Box[T any] struct { V T }
func (b Box[T]) Get() T { return b.V }
func main() { _ = Box[int]{V:1}.Get() }
