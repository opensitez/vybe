// vybe-test: go/blank_identifier_extended/blank_multi_assign_func_returns_compile
// origin: languages/go/tests/go/test_blank_identifier_extended.rs
// vybe-test-mode: compile

package main
func duo() (int, int) { return 1, 2 }
func main() { _, y := duo()
_ = y }
