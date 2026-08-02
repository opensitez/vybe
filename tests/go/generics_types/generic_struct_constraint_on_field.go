// vybe-test: go/generics_types/generic_struct_constraint_on_field
// origin: languages/go/tests/go/test_generics_types.rs
// vybe-test-mode: compile

package main
type Keyer interface { Key() string }
type Named struct { Name string }
func (n Named) Key() string { return n.Name }
type Entry[T Keyer] struct { Item T }
func main() { _ = Entry[Named]{Item: Named{Name: "a"}} }
