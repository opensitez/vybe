// vybe-test: go/generics_constraints_extended/generic_interface_constraint_two_methods
// origin: languages/go/tests/go/test_generics_constraints_extended.rs
// vybe-test-mode: compile

package main
type RW interface { Read() int
Write(int) }
func Use[T RW](v T) { v.Write(1)
_ = v.Read() }
type S struct { n int }
func (s *S) Read() int { return s.n }
func (s *S) Write(n int) { s.n = n }
func main() { var x S
Use(&x) }
