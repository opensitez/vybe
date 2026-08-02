// vybe-test: go/builtins_expressions_extra/complex_parts_in_assignment_compile
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs
// vybe-test-mode: compile

package main
func main() { r, i := real(complex(5, 6)), imag(complex(5, 6))
_, _ = r, i }
