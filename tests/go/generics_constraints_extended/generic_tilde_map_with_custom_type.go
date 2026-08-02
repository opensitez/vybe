// vybe-test: go/generics_constraints_extended/generic_tilde_map_with_custom_type
// origin: languages/go/tests/go/test_generics_constraints_extended.rs
// vybe-test-mode: compile

package main
type M map[string]int
func KeysLen[T ~map[string]int](m T) int { return len(m) }
func main() { _ = KeysLen(M{"a": 1}) }
