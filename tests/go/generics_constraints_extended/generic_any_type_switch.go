// vybe-test: go/generics_constraints_extended/generic_any_type_switch
// origin: languages/go/tests/go/test_generics_constraints_extended.rs
// vybe-test-mode: compile

package main
func Kind[T any](v T) string { switch any(v).(type) { case int: return "int"
case string: return "string"
default: return "other" } }
func main() { _ = Kind(1.0) }
