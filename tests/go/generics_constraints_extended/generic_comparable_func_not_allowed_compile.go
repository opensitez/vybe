// vybe-test: go/generics_constraints_extended/generic_comparable_func_not_allowed_compile
// origin: languages/go/tests/go/test_generics_constraints_extended.rs
// vybe-test-mode: compile

package main
func Bad[K comparable]() {}
func main() { type F func()
_ = F }
