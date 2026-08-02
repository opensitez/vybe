// vybe-test: go/generics_constraints_extended/generic_tilde_slice_to_custom
// origin: languages/go/tests/go/test_generics_constraints_extended.rs
// vybe-test-mode: compile

package main
type Names []string
func First[N ~[]string](n N) string { if len(n) == 0 { return "" }
return n[0] }
func main() { _ = First(Names{"go"}) }
