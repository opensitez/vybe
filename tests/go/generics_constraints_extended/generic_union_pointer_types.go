// vybe-test: go/generics_constraints_extended/generic_union_pointer_types
// origin: languages/go/tests/go/test_generics_constraints_extended.rs
// vybe-test-mode: compile

package main
func Deref[T *int | *string](p T) interface{} { switch v := any(p).(type) { case *int: return *v
default: return *v.(*string) } }
func main() { x := 1
_ = Deref(&x) }
