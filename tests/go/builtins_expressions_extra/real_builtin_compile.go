// vybe-test: go/builtins_expressions_extra/real_builtin_compile
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs
// vybe-test-mode: compile

package main
func main() { r := real(complex(3, 4))
_ = r }
