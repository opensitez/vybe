// vybe-test: go/generics_types/generic_interface_type_param_decl
// origin: languages/go/tests/go/test_generics_types.rs
// vybe-test-mode: compile

package main
type Storer[T any] interface { Store(T) }
func main() {}
