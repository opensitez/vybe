// vybe-test: go/generics_types/generic_multi_constraint_interface_embed
// origin: languages/go/tests/go/test_generics_types.rs
// vybe-test-mode: compile

package main
import "cmp"
type KeyedOrdered[T cmp.Ordered] interface { cmp.Ordered
Key() T }
func main() {}
