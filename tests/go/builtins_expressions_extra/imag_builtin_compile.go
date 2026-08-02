// vybe-test: go/builtins_expressions_extra/imag_builtin_compile
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs
// vybe-test-mode: compile

package main
func main() { i := imag(complex(3, 4))
_ = i }
