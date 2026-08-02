// vybe-test: go/generics_functions/generic_pointer_constraint
// origin: languages/go/tests/go/test_generics_functions.rs
// vybe-test-mode: compile

package main
func Zero[T any](p *T) { var z T
*p = z }
func main() { var x int
Zero(&x) }
