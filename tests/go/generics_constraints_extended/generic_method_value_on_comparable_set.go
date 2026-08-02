// vybe-test: go/generics_constraints_extended/generic_method_value_on_comparable_set
// origin: languages/go/tests/go/test_generics_constraints_extended.rs
// vybe-test-mode: compile

package main
type IDSet[T comparable] map[T]struct{}
func (s IDSet[T]) Add(v T) { s[v] = struct{}{} }
func main() { m := IDSet[int]{}
m.Add(1) }
